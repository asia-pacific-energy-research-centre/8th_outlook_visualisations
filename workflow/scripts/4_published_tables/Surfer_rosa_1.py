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
Fuels = EGEDA_years_reference.fuel_code.unique()
Items = EGEDA_years_reference.item_code_new.unique()

# Define colour palette

colours_dict = pd.read_csv('./data/2_Mapping_and_other/colours_dict.csv',\
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
Economy_codes = ['APEC']

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
    non_zero = (ref_fedfuel_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_fedfuel_1 = ref_fedfuel_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (ref_fedsector_2.loc[:,'2000':] != 0).any(axis = 1)
    ref_fedsector_2 = ref_fedsector_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (ref_bld_2.loc[:,'2000':] != 0).any(axis = 1)
    ref_bld_2 = ref_bld_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (ref_ind_2.loc[:,'2000':] != 0).any(axis = 1)
    ref_ind_2 = ref_ind_2.loc[non_zero].reset_index(drop = True)
    
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
    non_zero = (ref_trn_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_trn_1 = ref_trn_1.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (ref_ag_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_ag_1 = ref_ag_1.loc[non_zero].reset_index(drop = True)
    
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
    non_zero = (netz_fedfuel_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_fedfuel_1 = netz_fedfuel_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_fedsector_2.loc[:,'2000':] != 0).any(axis = 1)
    netz_fedsector_2 = netz_fedsector_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_bld_2.loc[:,'2000':] != 0).any(axis = 1)
    netz_bld_2 = netz_bld_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_ind_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_ind_1 = netz_ind_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_ind_2.loc[:,'2000':] != 0).any(axis = 1)
    netz_ind_2 = netz_ind_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_trn_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_trn_1 = netz_trn_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_trn_2.loc[:,'2018':] != 0).any(axis = 1)
    netz_trn_2 = netz_trn_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_ag_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_ag_1 = netz_ag_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_hyd_1.loc[:,'2018':] != 0).any(axis = 1)
    netz_hyd_1 = netz_hyd_1.loc[non_zero].reset_index(drop = True)

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

    ref_exports_temp1 = ref_exports_2.copy().select_dtypes(include = [np.number]) * -1
    ref_exports_temp2 = ref_exports_2.copy()
    ref_exports_temp2[ref_exports_temp1.columns] = ref_exports_temp1

    # Net trade

    ref_nettrade_1 = ref_imports_2.copy().append(ref_exports_temp2).groupby('fuel_code').sum()\
        .assign(item_code_new = 'Net trade').reset_index()

    ref_nettrade_1 = ref_nettrade_1[['fuel_code', 'item_code_new'] + col_chart_years]

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
    non_zero = (netz_tpes_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_tpes_1 = netz_tpes_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_prod_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_prod_1 = netz_prod_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_tpes_comp_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_tpes_comp_1 = netz_tpes_comp_1.loc[non_zero].reset_index(drop = True)

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

    netz_exports_temp1 = netz_exports_2.copy().select_dtypes(include = [np.number]) * -1
    netz_exports_temp2 = netz_exports_2.copy()
    netz_exports_temp2[netz_exports_temp1.columns] = netz_exports_temp1

    # Net trade

    netz_nettrade_1 = netz_imports_2.copy().append(netz_exports_temp2).groupby('fuel_code').sum()\
        .assign(item_code_new = 'Net trade').reset_index()

    netz_nettrade_1 = netz_nettrade_1[['fuel_code', 'item_code_new'] + col_chart_years]

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
    non_zero = (ref_pow_use_2.loc[:,'2000':] != 0).any(axis = 1)
    ref_pow_use_2 = ref_pow_use_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (ref_elecgen_2.loc[:,'2000':] != 0).any(axis = 1)
    ref_elecgen_2 = ref_elecgen_2.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (ref_refinery_1.loc[:,'2017':] != 0).any(axis = 1)
    ref_refinery_1 = ref_refinery_1.loc[non_zero].reset_index(drop = True)

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

    ref_powcap_1 = ref_powcap_1.append([coal_capacity, coal_ccs_capacity, gas_capacity, gas_ccs_capacity, oil_capacity, nuclear_capacity,
                                            hydro_capacity, bio_capacity, wind_capacity, solar_capacity, 
                                            storage_capacity, geo_capacity, waste_capacity, other_capacity])\
        [['TECHNOLOGY'] + list(ref_powcap_1.loc[:, '2018':'2050'])].reset_index(drop = True) 

    ref_powcap_1 = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(pow_capacity_agg)].reset_index(drop = True)

    ref_powcap_1['TECHNOLOGY'] = pd.Categorical(ref_powcap_1['TECHNOLOGY'], prod_agg_tech)

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
    non_zero = (ref_powcap_1.loc[:,'2018':] != 0).any(axis = 1)
    ref_powcap_1 = ref_powcap_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (ref_ownuse_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_ownuse_1 = ref_ownuse_1.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (ref_heatgen_2.loc[:,'2000':] != 0).any(axis = 1)
    ref_heatgen_2 = ref_heatgen_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_pow_use_2.loc[:,'2000':] != 0).any(axis = 1)
    netz_pow_use_2 = netz_pow_use_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_elecgen_2.loc[:,'2000':] != 0).any(axis = 1)
    netz_elecgen_2 = netz_elecgen_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_elecgen_4.loc[:,'2000':] != 0).any(axis = 1)
    netz_elecgen_4 = netz_elecgen_4.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (netz_refinery_1.loc[:,'2017':] != 0).any(axis = 1)
    netz_refinery_1 = netz_refinery_1.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (netz_refinery_2.loc[:,'2017':] != 0).any(axis = 1)
    netz_refinery_2 = netz_refinery_2.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (netz_hydrogen_2.loc[:,'2018':] != 0).any(axis = 1)
    netz_hydrogen_2 = netz_hydrogen_2.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (netz_hyd_use_1.loc[:,'2018':] != 0).any(axis = 1)
    netz_hyd_use_1 = netz_hyd_use_1.loc[non_zero].reset_index(drop = True)

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

    netz_powcap_1 = netz_powcap_1.append([coal_capacity, coal_ccs_capacity, gas_capacity, gas_ccs_capacity, oil_capacity, nuclear_capacity,
                                            hydro_capacity, bio_capacity, wind_capacity, solar_capacity, 
                                            storage_capacity, geo_capacity, waste_capacity, other_capacity])\
        [['TECHNOLOGY'] + list(netz_powcap_1.loc[:,'2018':'2050'])].reset_index(drop = True) 

    netz_powcap_1 = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(pow_capacity_agg)].reset_index(drop = True)

    netz_powcap_1['TECHNOLOGY'] = pd.Categorical(netz_powcap_1['TECHNOLOGY'], prod_agg_tech)

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
    non_zero = (netz_powcap_1.loc[:,'2018':] != 0).any(axis = 1)
    netz_powcap_1 = netz_powcap_1.loc[non_zero].reset_index(drop = True)

    netz_powcap_1_rows = netz_powcap_1.shape[0]
    netz_powcap_1_cols = netz_powcap_1.shape[1]

    netz_powcap_2 = netz_powcap_1[['TECHNOLOGY'] + trans_col_chart]

    netz_powcap_2_rows = netz_powcap_2.shape[0]
    netz_powcap_2_cols = netz_powcap_2.shape[1]

    # Get rid of zero rows
    non_zero = (netz_powcap_3.loc[:,'2018':] != 0).any(axis = 1)
    netz_powcap_3 = netz_powcap_3.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (netz_trans_3.loc[:,'2017':] != 0).any(axis = 1)
    netz_trans_3 = netz_trans_3.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_ownuse_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_ownuse_1 = netz_ownuse_1.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (netz_heatgen_2.loc[:,'2000':] != 0).any(axis = 1)
    netz_heatgen_2 = netz_heatgen_2.loc[non_zero].reset_index(drop = True)

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

    # Get rid of zero rows
    non_zero = (netz_heat_use_2.loc[:,'2017':] != 0).any(axis = 1)
    netz_heat_use_2 = netz_heat_use_2.loc[non_zero].reset_index(drop = True)

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

        if economy == 'APEC':
            target_row = ['APEC', 'Target'] + [55] * 51
            target_series = pd.Series(target_row, index = ref_enint_3.columns)

            ref_enint_3 = ref_enint_3.append(target_series, ignore_index = True).reset_index(drop = True)

        else:
            pass

        ref_enint_3_rows = ref_enint_3.shape[0]
        ref_enint_3_cols = ref_enint_3.shape[1]

        # CARBON NEUTRALITY

        netz_enint_1 = netz_tfec_1.copy()
        netz_enint_1['Economy'] = economy
        netz_enint_1['Series'] = 'TFEC'

        netz_enint_1 = netz_enint_1.append(macro_1[macro_1['Series'] == 'GDP 2018 USD PPP']).copy().reset_index(drop = True)

        netz_enint_1 = netz_enint_1[['Economy', 'Series'] + list(netz_enint_1.loc[:,'2000':'2050'])]

        netz_ei_calc1 = [economy, 'TFEC energy intensity'] + list(netz_enint_1.iloc[0, 2:] / netz_enint_1.iloc[1, 2:])
        netz_ei_series1 = pd.Series(netz_ei_calc1, index = netz_enint_1.columns)

        netz_enint_2 = netz_enint_1.append(netz_ei_series1, ignore_index = True).reset_index(drop = True)

        netz_ei_calc2 = [economy, 'Carbon Neutrality'] + list(netz_enint_2.iloc[2, 2:] / netz_enint_2.iloc[2, 7] * 100)
        netz_ei_series2 = pd.Series(netz_ei_calc2, index = netz_enint_2.columns)

        netz_enint_3 = netz_enint_2.append(netz_ei_series2, ignore_index = True).reset_index(drop = True)

        netz_enint_3.loc[netz_enint_3['Series'] == 'Carbon Neutrality', '2000':'2017'] = np.nan

        if economy == 'APEC':
            target_row2 = ['APEC', 'Target'] + [55] * 51
            target_series2 = pd.Series(target_row2, index = netz_enint_3.columns)

            netz_enint_3 = netz_enint_3.append(target_series2, ignore_index = True).reset_index(drop = True)

        else:
            pass

        netz_enint_3_rows = netz_enint_3.shape[0]
        netz_enint_3_cols = netz_enint_3.shape[1]

    else:
        ref_enint_3 = pd.DataFrame()
        ref_enint_3_rows = ref_enint_3.shape[0]
        ref_enint_3_cols = ref_enint_3.shape[1]

        netz_enint_3 = pd.DataFrame()
        netz_enint_3_rows = netz_enint_3.shape[0]
        netz_enint_3_cols = netz_enint_3.shape[1]

    # Energy supply intensity

    if any(economy in s for s in list(macro_GDP['Economy'])):

        # REFERENCE
        ref_enint_sup1 = ref_tpes_1[ref_tpes_1['fuel_code'] == 'Total'].copy().reset_index(drop = True)
        ref_enint_sup1['Economy'] = economy
        ref_enint_sup1['Series'] = 'TPES'

        ref_enint_sup1 = ref_enint_sup1.append(macro_1[macro_1['Series'] == 'GDP 2018 USD PPP']).copy().reset_index(drop = True)

        ref_enint_sup1 = ref_enint_sup1[['Economy', 'Series'] + list(ref_enint_sup1.loc[:, '2000':'2050'])]

        ref_calc1 = [economy, 'TPES energy intensity'] + list(ref_enint_sup1.iloc[0, 2:] / ref_enint_sup1.iloc[1, 2:])
        ref_series1 = pd.Series(ref_calc1, index = ref_enint_sup1.columns)

        ref_enint_sup2 = ref_enint_sup1.append(ref_series1, ignore_index = True).reset_index(drop = True)

        ref_calc2 = [economy, 'Reference'] + list(ref_enint_sup2.iloc[2, 2:] / ref_enint_sup2.iloc[2, 7] * 100)
        ref_series2 = pd.Series(ref_calc2, index = ref_enint_sup2.columns)

        ref_enint_sup3 = ref_enint_sup2.append(ref_series2, ignore_index = True).reset_index(drop = True)

        if economy == 'APEC':
            target_row = ['APEC', 'Target'] + [55] * 51
            target_series = pd.Series(target_row, index = ref_enint_sup3.columns)

            ref_enint_sup3 = ref_enint_sup3.append(target_series, ignore_index = True).reset_index(drop = True)

        else:
            pass

        ref_enint_sup3_rows = ref_enint_sup3.shape[0]
        ref_enint_sup3_cols = ref_enint_sup3.shape[1]

        # CARBON NEUTRALITY
        netz_enint_sup1 = netz_tpes_1[netz_tpes_1['fuel_code'] == 'Total'].copy().reset_index(drop = True)
        netz_enint_sup1['Economy'] = economy
        netz_enint_sup1['Series'] = 'TPES'

        netz_enint_sup1 = netz_enint_sup1.append(macro_1[macro_1['Series'] == 'GDP 2018 USD PPP']).copy().reset_index(drop = True)

        netz_enint_sup1 = netz_enint_sup1[['Economy', 'Series'] + list(netz_enint_sup1.loc[:, '2000':'2050'])]

        netz_calc1 = [economy, 'TPES energy intensity'] + list(netz_enint_sup1.iloc[0, 2:] / netz_enint_sup1.iloc[1, 2:])
        netz_series1 = pd.Series(netz_calc1, index = netz_enint_sup1.columns)

        netz_enint_sup2 = netz_enint_sup1.append(netz_series1, ignore_index = True).reset_index(drop = True)

        netz_calc2 = [economy, 'Carbon Neutrality'] + list(netz_enint_sup2.iloc[2, 2:] / netz_enint_sup2.iloc[2, 7] * 100)
        netz_series2 = pd.Series(netz_calc2, index = netz_enint_sup2.columns)

        netz_enint_sup3 = netz_enint_sup2.append(netz_series2, ignore_index = True).reset_index(drop = True)

        # Remove CN historical
        netz_enint_sup3.loc[netz_enint_sup3['Series'] == 'Carbon Neutrality', '2000':'2017'] = np.nan

        if economy == 'APEC':
            target_row2 = ['APEC', 'Target'] + [55] * 51
            target_series2 = pd.Series(target_row2, index = netz_enint_sup3.columns)

            netz_enint_sup3 = netz_enint_sup3.append(target_series2, ignore_index = True).reset_index(drop = True)

        else:
            pass

        netz_enint_sup3_rows = netz_enint_sup3.shape[0]
        netz_enint_sup3_cols = netz_enint_sup3.shape[1]

    else:
        ref_enint_sup3 = pd.DataFrame()
        ref_enint_sup3_rows = ref_enint_sup3.shape[0]
        ref_enint_sup3_cols = ref_enint_3.shape[1]

        netz_enint_sup3 = pd.DataFrame()
        netz_enint_sup3_rows = netz_enint_sup3.shape[0]
        netz_enint_sup3_cols = netz_enint_sup3.shape[1]

    ##############################################################################################################

    # OSeMOSYS datafrane builds

    # REFERENCE
    # Steel
    if any(economy in s for s in list(ref_steel_2['REGION'])):
        
        ref_steel_3 = ref_steel_2[ref_steel_2['REGION'] == economy].copy()\
            [['Industry', 'tech_mix'] + list(ref_steel_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        ref_steel_3_rows = ref_steel_3.shape[0]
        ref_steel_3_cols = ref_steel_3.shape[1]

    else:
        ref_steel_3 = pd.DataFrame()
        ref_steel_3_rows = ref_steel_3.shape[0]
        ref_steel_3_cols = ref_steel_3.shape[1]

    # Chemicals
    if any(economy in s for s in list(ref_chem_2['REGION'])):
        
        ref_chem_3 = ref_chem_2[ref_chem_2['REGION'] == economy].copy()\
            [['Industry', 'tech_mix'] + list(ref_chem_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        ref_chem_3_rows = ref_chem_3.shape[0]
        ref_chem_3_cols = ref_chem_3.shape[1]

    else:
        ref_chem_3 = pd.DataFrame()
        ref_chem_3_rows = ref_chem_3.shape[0]
        ref_chem_3_cols = ref_chem_3.shape[1]

    # Cement
    if any(economy in s for s in list(ref_cement_2['REGION'])):
        
        ref_cement_3 = ref_cement_2[ref_cement_2['REGION'] == economy].copy()\
            [['Industry', 'tech_mix'] + list(ref_cement_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        ref_cement_3_rows = ref_cement_3.shape[0]
        ref_cement_3_cols = ref_cement_3.shape[1]

    else:
        ref_cement_3 = pd.DataFrame()
        ref_cement_3_rows = ref_cement_3.shape[0]
        ref_cement_3_cols = ref_cement_3.shape[1]

    # CARBON NEUTRALITY
    # Steel
    if any(economy in s for s in list(netz_steel_2['REGION'])):
        
        netz_steel_3 = netz_steel_2[netz_steel_2['REGION'] == economy].copy()\
            [['Industry', 'tech_mix'] + list(netz_steel_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        netz_steel_3_rows = netz_steel_3.shape[0]
        netz_steel_3_cols = netz_steel_3.shape[1]

    else:
        netz_steel_3 = pd.DataFrame()
        netz_steel_3_rows = netz_steel_3.shape[0]
        netz_steel_3_cols = netz_steel_3.shape[1]

    # Chemicals
    if any(economy in s for s in list(netz_chem_2['REGION'])):
        
        netz_chem_3 = netz_chem_2[netz_chem_2['REGION'] == economy].copy()\
            [['Industry', 'tech_mix'] + list(netz_chem_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        netz_chem_3_rows = netz_chem_3.shape[0]
        netz_chem_3_cols = netz_chem_3.shape[1]

    else:
        netz_chem_3 = pd.DataFrame()
        netz_chem_3_rows = netz_chem_3.shape[0]
        netz_chem_3_cols = netz_chem_3.shape[1]

    # Cement
    if any(economy in s for s in list(netz_cement_2['REGION'])):
        
        netz_cement_3 = netz_cement_2[netz_cement_2['REGION'] == economy].copy()\
            [['Industry', 'tech_mix'] + list(netz_cement_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        netz_cement_3_rows = netz_cement_3.shape[0]
        netz_cement_3_cols = netz_cement_3.shape[1]

    else:
        netz_cement_3 = pd.DataFrame()
        netz_cement_3_rows = netz_cement_3.shape[0]
        netz_cement_3_cols = netz_cement_3.shape[1]

    # TRANSPORT REFERENCE
    # Road modality 
    if any(economy in s for s in list(ref_roadmodal_2['REGION'])):

        ref_roadmodal_3 = ref_roadmodal_2[ref_roadmodal_2['REGION'] == economy].copy()\
            [['Transport', 'modality'] + list(ref_roadmodal_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        ref_roadmodal_3_rows = ref_roadmodal_3.shape[0]
        ref_roadmodal_3_cols = ref_roadmodal_3.shape[1]

    else:
        ref_roadmodal_3 = pd.DataFrame()
        ref_roadmodal_3_rows = ref_roadmodal_3.shape[0]
        ref_roadmodal_3_cols = ref_roadmodal_3.shape[1]

    # Fuel modality 
    if any(economy in s for s in list(ref_roadfuel_2['REGION'])):

        ref_roadfuel_3 = ref_roadfuel_2[ref_roadfuel_2['REGION'] == economy].copy()\
            [['Transport', 'modality'] + list(ref_roadfuel_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        ref_roadfuel_3_rows = ref_roadfuel_3.shape[0]
        ref_roadfuel_3_cols = ref_roadfuel_3.shape[1]

    else:
        ref_roadfuel_3 = pd.DataFrame()
        ref_roadfuel_3_rows = ref_roadfuel_3.shape[0]
        ref_roadfuel_3_cols = ref_roadfuel_3.shape[1]

    # TRANSPORT CARBON NEUTRALITY
    # Road modality 
    if any(economy in s for s in list(netz_roadmodal_2['REGION'])):

        netz_roadmodal_3 = netz_roadmodal_2[netz_roadmodal_2['REGION'] == economy].copy()\
            [['Transport', 'modality'] + list(netz_roadmodal_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        netz_roadmodal_3_rows = netz_roadmodal_3.shape[0]
        netz_roadmodal_3_cols = netz_roadmodal_3.shape[1]

    else:
        netz_roadmodal_3 = pd.DataFrame()
        netz_roadmodal_3_rows = netz_roadmodal_3.shape[0]
        netz_roadmodal_3_cols = netz_roadmodal_3.shape[1]

    # Fuel modality 
    if any(economy in s for s in list(netz_roadfuel_2['REGION'])):

        netz_roadfuel_3 = netz_roadfuel_2[netz_roadfuel_2['REGION'] == economy].copy()\
            [['Transport', 'modality'] + list(netz_roadfuel_2.loc[:, '2018':'2050'])].reset_index(drop = True)

        netz_roadfuel_3_rows = netz_roadfuel_3.shape[0]
        netz_roadfuel_3_cols = netz_roadfuel_3.shape[1]

    else:
        netz_roadfuel_3 = pd.DataFrame()
        netz_roadfuel_3_rows = netz_roadfuel_3.shape[0]
        netz_roadfuel_3_cols = netz_roadfuel_3.shape[1]


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
    non_zero = (ref_emiss_fuel_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_emiss_fuel_1 = ref_emiss_fuel_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (ref_emiss_sector_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_emiss_sector_1 = ref_emiss_sector_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_emiss_fuel_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_emiss_fuel_1 = netz_emiss_fuel_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_emiss_sector_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_emiss_sector_1 = netz_emiss_sector_1.loc[non_zero].reset_index(drop = True)

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

    # New emissions dataframe (for wedge chart)

    # SECTOR

    # Carbon neutrality emissions minus Reference emissions

    emiss_diff_sector = netz_emiss_sector_1.iloc[:-1,:].select_dtypes(include = [np.number]) - \
        ref_emiss_sector_1.iloc[:-1,:].select_dtypes(include = [np.number])

    emissions_wedge_1 = ref_emiss_sector_1.iloc[:-1,:].copy()

    emissions_wedge_1[emiss_diff_sector.columns] = emiss_diff_sector

    # Now add phantom row for wedge

    emissions_wedge_1 = emissions_wedge_1.append(emiss_total_1.iloc[0,:].copy()).reset_index(drop = True)

    emissions_wedge_1.loc[emissions_wedge_1['fuel_code'] == 'Reference', 'fuel_code'] = np.nan
    emissions_wedge_1.loc[emissions_wedge_1['item_code_new'] == 'Emissions', 'item_code_new'] = np.nan

    emissions_wedge_1 = emissions_wedge_1.append(emiss_total_1.copy()).reset_index(drop = True)

    emissions_wedge_1.loc[emissions_wedge_1['fuel_code'] == 'Reference', 'item_code_new'] = 'Reference'
    emissions_wedge_1.loc[emissions_wedge_1['fuel_code'] == 'Carbon Neutrality', 'item_code_new'] = 'Carbon Neutrality'
    emissions_wedge_1.loc[emissions_wedge_1['item_code_new'] == 'Reference', 'fuel_code'] = '19_total'
    emissions_wedge_1.loc[emissions_wedge_1['item_code_new'] == 'Carbon Neutrality', 'fuel_code'] = '19_total'

    # Get rid of data for CN for historical
    emissions_wedge_1.loc[emissions_wedge_1['item_code_new'] == 'Carbon Neutrality', '2000':'2017'] = np.nan

    emissions_wedge_1_rows = emissions_wedge_1.shape[0]
    emissions_wedge_1_cols = emissions_wedge_1.shape[1]

    # FUEL

    # Carbon neutrality emissions minus Reference emissions

    emiss_diff_fuel = netz_emiss_fuel_1.iloc[:-1,:].select_dtypes(include = [np.number]) - \
        ref_emiss_fuel_1.iloc[:-1,:].select_dtypes(include = [np.number])

    emissions_wedge_2 = ref_emiss_fuel_1.iloc[:-1,:].copy()

    emissions_wedge_2[emiss_diff_fuel.columns] = emiss_diff_fuel

    # Now add phantom row for wedge

    emissions_wedge_2 = emissions_wedge_2.append(emiss_total_1.iloc[0,:].copy()).reset_index(drop = True)

    emissions_wedge_2.loc[emissions_wedge_2['fuel_code'] == 'Reference', 'fuel_code'] = np.nan
    emissions_wedge_2.loc[emissions_wedge_2['item_code_new'] == 'Emissions', 'item_code_new'] = np.nan

    emissions_wedge_2 = emissions_wedge_2.append(emiss_total_1.copy()).reset_index(drop = True)

    emissions_wedge_2.loc[emissions_wedge_2['fuel_code'] == 'Carbon Neutrality', 'fuel_code'] = 'Carbon Neutrality'

    # Get rid of data for CN for historical
    emissions_wedge_2.loc[emissions_wedge_2['fuel_code'] == 'Carbon Neutrality', '2000':'2017'] = np.nan

    emissions_wedge_2_rows = emissions_wedge_2.shape[0]
    emissions_wedge_2_cols = emissions_wedge_2.shape[1]


    ##################################################################################################

    # Fuel dataframe builds

    # ref_coal_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
    #                                    (EGEDA_years_reference['item_code_new'].isin(no_trad_bio_sectors)) &
    #                                    (EGEDA_years_reference['fuel_code'] == '1_coal')]\
    #                                        .loc[:, 'fuel_code':].reset_index(drop = True)

    # REFERENCE

    # Coal
    ref_coal_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                       (EGEDA_years_reference['item_code_new'].isin(fuel_vector_1)) &
                                       (EGEDA_years_reference['fuel_code'].isin(['1_coal', '2_coal_products']))]\
                                           .copy().groupby(['item_code_new']).sum().assign(fuel_code = 'Coal').reset_index()\
                                           [['fuel_code', 'item_code_new'] + col_chart_years] 

    ref_coal_1.loc[ref_coal_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_coal_1.loc[ref_coal_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_coal_1.loc[ref_coal_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    ref_coal_1.loc[ref_coal_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    ref_coal_1.loc[ref_coal_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    ref_coal_1 = ref_coal_1[ref_coal_1['item_code_new'].isin(fuel_final_nobunk)].reset_index(drop = True)

    ref_coal_1['item_code_new'] = pd.Categorical(
        ref_coal_1['item_code_new'], 
        categories = fuel_final_nobunk, 
        ordered = True)

    ref_coal_1 = ref_coal_1.sort_values('item_code_new').reset_index(drop = True)

    ref_coal_1_rows = ref_coal_1.shape[0]
    ref_coal_1_cols = ref_coal_1.shape[1]

    # split into thermal and metallurgical

    ref_coal_2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                       (EGEDA_years_reference['item_code_new'].isin(fuel_vector_1)) &
                                       (EGEDA_years_reference['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                            '1_x_coal_thermal', '2_coal_products']))]\
                                                .copy().reset_index(drop = True)

    met_coal = ref_coal_2[ref_coal_2['fuel_code'].isin(['1_1_coking_coal', '2_coal_products'])].copy()\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Metallurgical coal').reset_index()

    ref_coaltype_1 = ref_coal_2.append(met_coal).reset_index(drop = True)

    ref_coaltype_1.loc[ref_coaltype_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_coaltype_1.loc[ref_coaltype_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_coaltype_1.loc[ref_coaltype_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    ref_coaltype_1.loc[ref_coaltype_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    ref_coaltype_1.loc[ref_coaltype_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'
    ref_coaltype_1.loc[ref_coaltype_1['fuel_code'] == '1_x_coal_thermal', 'fuel_code'] = 'Thermal coal'
    ref_coaltype_1.loc[ref_coaltype_1['fuel_code'] == '1_5_lignite', 'fuel_code'] = 'Lignite'

    ref_coaltype_1 = ref_coaltype_1[ref_coaltype_1['item_code_new'].isin(fuel_final_nobunk)].reset_index(drop = True)
    ref_coaltype_1 = ref_coaltype_1[ref_coaltype_1['fuel_code'].isin(['Thermal coal', 'Lignite', 'Metallurgical coal'])].reset_index(drop = True)

    ref_coaltype_1['item_code_new'] = pd.Categorical(
        ref_coaltype_1['item_code_new'], 
        categories = fuel_final_nobunk, 
        ordered = True)

    ref_coaltype_1 = ref_coaltype_1.sort_values('item_code_new').reset_index(drop = True)

    ref_coaltype_1['fuel_code'] = pd.Categorical(
        ref_coaltype_1['fuel_code'], 
        categories = ['Thermal coal', 'Lignite', 'Metallurgical coal'], 
        ordered = True)

    ref_coaltype_1 = ref_coaltype_1.sort_values('fuel_code').reset_index(drop = True)

    ref_coaltype_1 = ref_coaltype_1[['fuel_code', 'item_code_new'] + list(ref_coaltype_1.loc[:,'2000':'2050'])]\
        .replace(np.nan, 0)

    # Get rid of zero rows
    non_zero = (ref_coaltype_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_coaltype_1 = ref_coaltype_1.loc[non_zero].reset_index(drop = True)

    ref_ct_prod1 = ref_coaltype_1[ref_coaltype_1['item_code_new'] == 'Production'].copy().reset_index(drop = True)

    ref_ct_prod1_rows = ref_ct_prod1.shape[0]
    ref_ct_prod1_cols = ref_ct_prod1.shape[1]

    ref_ct_imports1 = ref_coaltype_1[ref_coaltype_1['item_code_new'] == 'Imports'].copy().reset_index(drop = True)

    ref_ct_imports1_rows = ref_ct_imports1.shape[0]
    ref_ct_imports1_cols = ref_ct_imports1.shape[1]

    ref_ct_exports1 = ref_coaltype_1[ref_coaltype_1['item_code_new'] == 'Exports'].copy().reset_index(drop = True)

    neg_to_pos = ref_ct_exports1.select_dtypes(include = [np.number]) * -1  
    ref_ct_exports1[neg_to_pos.columns] = neg_to_pos

    ref_ct_exports1_rows = ref_ct_exports1.shape[0]
    ref_ct_exports1_cols = ref_ct_exports1.shape[1]

    # Crude

    ref_crude_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(fuel_vector_1)) &
                                        (EGEDA_years_reference['fuel_code'] == '6_crude_oil_and_ngl')].copy()\
                                            [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                .reset_index(drop = True)
    
    ref_crude_1['fuel_code'].replace({'6_crude_oil_and_ngl': 'Crude oil and NGL'}, inplace=True)

    ref_crude_1.loc[ref_crude_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_crude_1.loc[ref_crude_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_crude_1.loc[ref_crude_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    ref_crude_1.loc[ref_crude_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    ref_crude_1.loc[ref_crude_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    ref_crude_1 = ref_crude_1[ref_crude_1['item_code_new'].isin(fuel_final_nobunk)].reset_index(drop = True)

    ref_crude_1['item_code_new'] = pd.Categorical(
        ref_crude_1['item_code_new'], 
        categories = fuel_final_nobunk, 
        ordered = True)

    ref_crude_1 = ref_crude_1.sort_values('item_code_new').reset_index(drop = True)

    ref_crude_1_rows = ref_crude_1.shape[0]
    ref_crude_1_cols = ref_crude_1.shape[1]

    # Petprod moved below crudecons

    ref_gas_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(fuel_vector_1)) &
                                        (EGEDA_years_reference['fuel_code'] == '8_gas')].copy()\
                                            [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                .reset_index(drop = True)

    gas_bunkers = ref_gas_1[ref_gas_1['item_code_new'].isin(['4_international_marine_bunkers',
                                                             '5_international_aviation_bunkers'])]\
                                                                 .groupby(['fuel_code']).sum().assign(fuel_code = '8_gas', item_code_new = 'Bunkers')

    ref_gas_1 = ref_gas_1.append([gas_bunkers]).reset_index(drop = True)
    
    ref_gas_1['fuel_code'].replace({'8_gas': 'Gas'}, inplace = True)

    ref_gas_1.loc[ref_gas_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_gas_1.loc[ref_gas_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_gas_1.loc[ref_gas_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    ref_gas_1.loc[ref_gas_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    ref_gas_1.loc[ref_gas_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    ref_gas_1 = ref_gas_1[ref_gas_1['item_code_new'].isin(fuel_final_bunk)].reset_index(drop = True)

    ref_gas_1['item_code_new'] = pd.Categorical(
        ref_gas_1['item_code_new'], 
        categories = fuel_final_bunk, 
        ordered = True)

    ref_gas_1 = ref_gas_1.sort_values('item_code_new').reset_index(drop = True)

    ref_gas_1_rows = ref_gas_1.shape[0]
    ref_gas_1_cols = ref_gas_1.shape[1]

    # LNG and pipe
    # Imports
    ref_gasim_1 = ref_gastrade_df1[(ref_gastrade_df1['REGION'] == economy) &
                                   (ref_gastrade_df1['TECHNOLOGY'].str.contains('import'))].copy()\
                                       .rename(columns = {'TECHNOLOGY': 'Imports'})\
                                           [['Imports'] + list(ref_gastrade_df1.loc[:, '2018': '2050'])]\
                                               .reset_index(drop = True)

    ref_gasim_1.loc[ref_gasim_1['Imports'] == 'SUP_8_1_natural_gas_import', 'Imports'] = 'Pipeline'
    ref_gasim_1.loc[ref_gasim_1['Imports'] == 'SUP_8_2_lng_import', 'Imports'] = 'LNG'

    ref_gasim_1_rows = ref_gasim_1.shape[0]
    ref_gasim_1_cols = ref_gasim_1.shape[1]
    
    # Exports
    ref_gasex_1 = ref_gastrade_df1[(ref_gastrade_df1['REGION'] == economy) &
                                   (ref_gastrade_df1['TECHNOLOGY'].str.contains('export'))].copy()\
                                       .rename(columns = {'TECHNOLOGY': 'Exports'})\
                                           [['Exports'] + list(ref_gastrade_df1.loc[:, '2018': '2050'])]\
                                               .reset_index(drop = True)

    ref_gasex_1.loc[ref_gasex_1['Exports'] == 'SUP_8_1_natural_gas_export', 'Exports'] = 'Pipeline'
    ref_gasex_1.loc[ref_gasex_1['Exports'] == 'SUP_8_2_lng_export', 'Exports'] = 'LNG'

    ref_gasex_1_rows = ref_gasex_1.shape[0]
    ref_gasex_1_cols = ref_gasex_1.shape[1]

    # Nuclear 

    ref_nuke_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(fuel_vector_1)) &
                                        (EGEDA_years_reference['fuel_code'] == '9_nuclear')].copy()\
                                            [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                .reset_index(drop = True)
    
    ref_nuke_1['fuel_code'].replace({'9_nuclear': 'Nuclear'}, inplace=True)

    ref_nuke_1.loc[ref_nuke_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_nuke_1.loc[ref_nuke_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_nuke_1.loc[ref_nuke_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    ref_nuke_1.loc[ref_nuke_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    ref_nuke_1.loc[ref_nuke_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    ref_nuke_1 = ref_nuke_1[ref_nuke_1['item_code_new'].isin(['Production'])].reset_index(drop = True)

    ref_nuke_1_rows = ref_nuke_1.shape[0]
    ref_nuke_1_cols = ref_nuke_1.shape[1]

    ref_biomass_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                          (EGEDA_years_reference['item_code_new'].isin(fuel_vector_1)) &
                                          (EGEDA_years_reference['fuel_code'] == '15_solid_biomass')].copy()\
                                              [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                  .reset_index(drop = True)
    
    ref_biomass_1['fuel_code'].replace({'15_solid_biomass': 'Biomass'}, inplace=True)

    ref_biomass_1.loc[ref_biomass_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_biomass_1.loc[ref_biomass_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_biomass_1.loc[ref_biomass_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    ref_biomass_1.loc[ref_biomass_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    ref_biomass_1.loc[ref_biomass_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    ref_biomass_1 = ref_biomass_1[ref_biomass_1['item_code_new'].isin(fuel_final_nobunk)].reset_index(drop = True)

    ref_biomass_1['item_code_new'] = pd.Categorical(
        ref_biomass_1['item_code_new'], 
        categories = fuel_final_nobunk, 
        ordered = True)

    ref_biomass_1 = ref_biomass_1.sort_values('item_code_new').reset_index(drop = True)

    ref_biomass_1_rows = ref_biomass_1.shape[0]
    ref_biomass_1_cols = ref_biomass_1.shape[1]

    ref_biofuel_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                          (EGEDA_years_reference['item_code_new'].isin(fuel_vector_1)) &
                                          (EGEDA_years_reference['fuel_code'].isin(['16_1_biogas', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))]\
                                              .copy().groupby(['item_code_new']).sum().assign(fuel_code = 'Biofuels').reset_index()\
                                              [['fuel_code', 'item_code_new'] + col_chart_years]

    biofuel_bunkers = ref_biofuel_1[ref_biofuel_1['item_code_new'].isin(['4_international_marine_bunkers',
                                                                         '5_international_aviation_bunkers'])]\
                                                                             .groupby(['fuel_code']).sum().assign(fuel_code = 'Biofuels',
                                                                                                                  item_code_new = 'Bunkers')

    ref_biofuel_2 = ref_biofuel_1.append([biofuel_bunkers]).reset_index(drop = True)

    ref_biofuel_2.loc[ref_biofuel_2['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_biofuel_2.loc[ref_biofuel_2['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_biofuel_2.loc[ref_biofuel_2['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    ref_biofuel_2.loc[ref_biofuel_2['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    ref_biofuel_2.loc[ref_biofuel_2['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    ref_biofuel_2 = ref_biofuel_2[ref_biofuel_2['item_code_new'].isin(fuel_final_bunk)].reset_index(drop = True)

    ref_biofuel_2['item_code_new'] = pd.Categorical(
        ref_biofuel_2['item_code_new'], 
        categories = fuel_final_bunk, 
        ordered = True)

    ref_biofuel_2 = ref_biofuel_2.sort_values('item_code_new').reset_index(drop = True)

    ref_biofuel_2_rows = ref_biofuel_2.shape[0]
    ref_biofuel_2_cols = ref_biofuel_2.shape[1]

    # liquid and solid renewables

    ref_renew_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(fuel_vector_1)) &
                                        (EGEDA_years_reference['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable',
                                                                                  '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                  '16_8_other_liquid_biofuels']))]\
                                           .copy().groupby(['item_code_new']).sum().assign(fuel_code = 'Liquid and solid renewables').reset_index()\
                                           [['fuel_code', 'item_code_new'] + col_chart_years] 

    renew_bunkers = ref_renew_1[ref_renew_1['item_code_new'].isin(['4_international_marine_bunkers',
                                                                         '5_international_aviation_bunkers'])]\
                                                                             .groupby(['fuel_code']).sum().assign(fuel_code = 'Liquid and solid renewables',
                                                                                                                  item_code_new = 'Bunkers')

    ref_renew_2 = ref_renew_1.append([renew_bunkers]).reset_index(drop = True)

    ref_renew_2.loc[ref_renew_2['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_renew_2.loc[ref_renew_2['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_renew_2.loc[ref_renew_2['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    ref_renew_2.loc[ref_renew_2['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    ref_renew_2.loc[ref_renew_2['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    ref_renew_2 = ref_renew_2[ref_renew_2['item_code_new'].isin(fuel_final_bunk)].reset_index(drop = True)

    ref_renew_2['item_code_new'] = pd.Categorical(
        ref_renew_2['item_code_new'], 
        categories = fuel_final_bunk, 
        ordered = True)

    ref_renew_2 = ref_renew_2.sort_values('item_code_new').reset_index(drop = True)

    ref_renew_2_rows = ref_renew_2.shape[0]
    ref_renew_2_cols = ref_renew_2.shape[1]

    # CARBON NEUTRALITY

    netz_coal_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                       (EGEDA_years_netzero['item_code_new'].isin(fuel_vector_1)) &
                                       (EGEDA_years_netzero['fuel_code'].isin(['1_coal', '2_coal_products']))]\
                                           .copy().groupby(['item_code_new']).sum().assign(fuel_code = 'Coal').reset_index()\
                                           [['fuel_code', 'item_code_new'] + col_chart_years] 

    netz_coal_1.loc[netz_coal_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_coal_1.loc[netz_coal_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_coal_1.loc[netz_coal_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    netz_coal_1.loc[netz_coal_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    netz_coal_1.loc[netz_coal_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    netz_coal_1 = netz_coal_1[netz_coal_1['item_code_new'].isin(fuel_final_nobunk)].reset_index(drop = True)

    netz_coal_1['item_code_new'] = pd.Categorical(
        netz_coal_1['item_code_new'], 
        categories = fuel_final_nobunk, 
        ordered = True)

    netz_coal_1 = netz_coal_1.sort_values('item_code_new').reset_index(drop = True)

    netz_coal_1_rows = netz_coal_1.shape[0]
    netz_coal_1_cols = netz_coal_1.shape[1]

    # split into thermal and metallurgical

    netz_coal_2 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                       (EGEDA_years_netzero['item_code_new'].isin(fuel_vector_1)) &
                                       (EGEDA_years_netzero['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                            '1_x_coal_thermal', '2_coal_products']))]\
                                                .copy().reset_index(drop = True)

    met_coal = netz_coal_2[netz_coal_2['fuel_code'].isin(['1_1_coking_coal', '2_coal_products'])].copy()\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Metallurgical coal').reset_index()

    netz_coaltype_1 = netz_coal_2.append(met_coal).reset_index(drop = True)

    netz_coaltype_1.loc[netz_coaltype_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_coaltype_1.loc[netz_coaltype_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_coaltype_1.loc[netz_coaltype_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    netz_coaltype_1.loc[netz_coaltype_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    netz_coaltype_1.loc[netz_coaltype_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'
    netz_coaltype_1.loc[netz_coaltype_1['fuel_code'] == '1_x_coal_thermal', 'fuel_code'] = 'Thermal coal'
    netz_coaltype_1.loc[netz_coaltype_1['fuel_code'] == '1_5_lignite', 'fuel_code'] = 'Lignite'

    netz_coaltype_1 = netz_coaltype_1[netz_coaltype_1['item_code_new'].isin(fuel_final_nobunk)].reset_index(drop = True)
    netz_coaltype_1 = netz_coaltype_1[netz_coaltype_1['fuel_code'].isin(['Thermal coal', 'Lignite', 'Metallurgical coal'])].reset_index(drop = True)

    netz_coaltype_1['item_code_new'] = pd.Categorical(
        netz_coaltype_1['item_code_new'], 
        categories = fuel_final_nobunk, 
        ordered = True)

    netz_coaltype_1 = netz_coaltype_1.sort_values('item_code_new').reset_index(drop = True)

    netz_coaltype_1['fuel_code'] = pd.Categorical(
        netz_coaltype_1['fuel_code'], 
        categories = ['Thermal coal', 'Lignite', 'Metallurgical coal'], 
        ordered = True)

    netz_coaltype_1 = netz_coaltype_1.sort_values('fuel_code').reset_index(drop = True)

    netz_coaltype_1 = netz_coaltype_1[['fuel_code', 'item_code_new'] + list(netz_coaltype_1.loc[:,'2000':'2050'])]\
        .replace(np.nan, 0)

    # Get rid of zero rows
    non_zero = (netz_coaltype_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_coaltype_1 = netz_coaltype_1.loc[non_zero].reset_index(drop = True)

    netz_ct_prod1 = netz_coaltype_1[netz_coaltype_1['item_code_new'] == 'Production'].copy().reset_index(drop = True)

    netz_ct_prod1_rows = netz_ct_prod1.shape[0]
    netz_ct_prod1_cols = netz_ct_prod1.shape[1]

    netz_ct_imports1 = netz_coaltype_1[netz_coaltype_1['item_code_new'] == 'Imports'].copy().reset_index(drop = True)

    netz_ct_imports1_rows = netz_ct_imports1.shape[0]
    netz_ct_imports1_cols = netz_ct_imports1.shape[1]

    netz_ct_exports1 = netz_coaltype_1[netz_coaltype_1['item_code_new'] == 'Exports'].copy().reset_index(drop = True)

    neg_to_pos = netz_ct_exports1.select_dtypes(include = [np.number]) * -1  
    netz_ct_exports1[neg_to_pos.columns] = neg_to_pos

    netz_ct_exports1_rows = netz_ct_exports1.shape[0]
    netz_ct_exports1_cols = netz_ct_exports1.shape[1]

    # Crude
    netz_crude_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(fuel_vector_1)) &
                                        (EGEDA_years_netzero['fuel_code'] == '6_crude_oil_and_ngl')].copy()\
                                            [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                .reset_index(drop = True)
    
    netz_crude_1['fuel_code'].replace({'6_crude_oil_and_ngl': 'Crude oil and NGL'}, inplace=True)

    netz_crude_1.loc[netz_crude_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_crude_1.loc[netz_crude_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_crude_1.loc[netz_crude_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    netz_crude_1.loc[netz_crude_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    netz_crude_1.loc[netz_crude_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    netz_crude_1 = netz_crude_1[netz_crude_1['item_code_new'].isin(fuel_final_nobunk)].reset_index(drop = True)

    netz_crude_1['item_code_new'] = pd.Categorical(
        netz_crude_1['item_code_new'], 
        categories = fuel_final_nobunk, 
        ordered = True)

    netz_crude_1 = netz_crude_1.sort_values('item_code_new').reset_index(drop = True)

    netz_crude_1_rows = netz_crude_1.shape[0]
    netz_crude_1_cols = netz_crude_1.shape[1]

    ## Petprod moved below crudecons

    netz_gas_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(fuel_vector_1)) &
                                        (EGEDA_years_netzero['fuel_code'] == '8_gas')].copy()\
                                            [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                .reset_index(drop = True)

    gas_bunkers = netz_gas_1[netz_gas_1['item_code_new'].isin(['4_international_marine_bunkers',
                                                             '5_international_aviation_bunkers'])]\
                                                                 .groupby(['fuel_code']).sum().assign(fuel_code = '8_gas', item_code_new = 'Bunkers')

    netz_gas_1 = netz_gas_1.append([gas_bunkers]).reset_index(drop = True)
    
    netz_gas_1['fuel_code'].replace({'8_gas': 'Gas'}, inplace=True)

    netz_gas_1.loc[netz_gas_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_gas_1.loc[netz_gas_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_gas_1.loc[netz_gas_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    netz_gas_1.loc[netz_gas_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    netz_gas_1.loc[netz_gas_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    netz_gas_1 = netz_gas_1[netz_gas_1['item_code_new'].isin(fuel_final_bunk)].reset_index(drop = True)

    netz_gas_1['item_code_new'] = pd.Categorical(
        netz_gas_1['item_code_new'], 
        categories = fuel_final_bunk, 
        ordered = True)

    netz_gas_1 = netz_gas_1.sort_values('item_code_new').reset_index(drop = True)

    netz_gas_1_rows = netz_gas_1.shape[0]
    netz_gas_1_cols = netz_gas_1.shape[1]

    # LNG and pipe
    # Imports
    netz_gasim_1 = netz_gastrade_df1[(netz_gastrade_df1['REGION'] == economy) &
                                   (netz_gastrade_df1['TECHNOLOGY'].str.contains('import'))].copy()\
                                       .rename(columns = {'TECHNOLOGY': 'Imports'})\
                                           [['Imports'] + list(netz_gastrade_df1.loc[:, '2018': '2050'])]\
                                               .reset_index(drop = True)

    netz_gasim_1.loc[netz_gasim_1['Imports'] == 'SUP_8_1_natural_gas_import', 'Imports'] = 'Pipeline'
    netz_gasim_1.loc[netz_gasim_1['Imports'] == 'SUP_8_2_lng_import', 'Imports'] = 'LNG'

    netz_gasim_1_rows = netz_gasim_1.shape[0]
    netz_gasim_1_cols = netz_gasim_1.shape[1]
    
    # Exports
    netz_gasex_1 = netz_gastrade_df1[(netz_gastrade_df1['REGION'] == economy) &
                                   (netz_gastrade_df1['TECHNOLOGY'].str.contains('export'))].copy()\
                                       .rename(columns = {'TECHNOLOGY': 'Exports'})\
                                           [['Exports'] + list(netz_gastrade_df1.loc[:, '2018': '2050'])]\
                                               .reset_index(drop = True)

    netz_gasex_1.loc[netz_gasex_1['Exports'] == 'SUP_8_1_natural_gas_export', 'Exports'] = 'Pipeline'
    netz_gasex_1.loc[netz_gasex_1['Exports'] == 'SUP_8_2_lng_export', 'Exports'] = 'LNG'

    netz_gasex_1_rows = netz_gasex_1.shape[0]
    netz_gasex_1_cols = netz_gasex_1.shape[1]

    # Nuclear

    netz_nuke_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(fuel_vector_1)) &
                                        (EGEDA_years_netzero['fuel_code'] == '9_nuclear')].copy()\
                                            [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                .reset_index(drop = True)
    
    netz_nuke_1['fuel_code'].replace({'9_nuclear': 'Nuclear'}, inplace=True)

    netz_nuke_1.loc[netz_nuke_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_nuke_1.loc[netz_nuke_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_nuke_1.loc[netz_nuke_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    netz_nuke_1.loc[netz_nuke_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    netz_nuke_1.loc[netz_nuke_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    netz_nuke_1 = netz_nuke_1[netz_nuke_1['item_code_new'].isin(['Production'])].reset_index(drop = True)

    netz_nuke_1_rows = netz_nuke_1.shape[0]
    netz_nuke_1_cols = netz_nuke_1.shape[1]

    netz_biomass_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                          (EGEDA_years_netzero['item_code_new'].isin(fuel_vector_1)) &
                                          (EGEDA_years_netzero['fuel_code'] == '15_solid_biomass')].copy()\
                                              [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                  .reset_index(drop = True)
    
    netz_biomass_1['fuel_code'].replace({'15_solid_biomass': 'Biomass'}, inplace=True)

    netz_biomass_1.loc[netz_biomass_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_biomass_1.loc[netz_biomass_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_biomass_1.loc[netz_biomass_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    netz_biomass_1.loc[netz_biomass_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    netz_biomass_1.loc[netz_biomass_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    netz_biomass_1 = netz_biomass_1[netz_biomass_1['item_code_new'].isin(fuel_final_nobunk)].reset_index(drop = True)

    netz_biomass_1['item_code_new'] = pd.Categorical(
        netz_biomass_1['item_code_new'], 
        categories = fuel_final_nobunk, 
        ordered = True)

    netz_biomass_1 = netz_biomass_1.sort_values('item_code_new').reset_index(drop = True)

    netz_biomass_1_rows = netz_biomass_1.shape[0]
    netz_biomass_1_cols = netz_biomass_1.shape[1]

    netz_biofuel_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                          (EGEDA_years_netzero['item_code_new'].isin(fuel_vector_1)) &
                                          (EGEDA_years_netzero['fuel_code'].isin(['16_1_biogas', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))]\
                                              .copy().groupby(['item_code_new']).sum().assign(fuel_code = 'Biofuels').reset_index()\
                                              [['fuel_code', 'item_code_new'] + col_chart_years]

    biofuel_bunkers = netz_biofuel_1[netz_biofuel_1['item_code_new'].isin(['4_international_marine_bunkers',
                                                                         '5_international_aviation_bunkers'])]\
                                                                             .groupby(['fuel_code']).sum().assign(fuel_code = 'Biofuels',
                                                                                                                  item_code_new = 'Bunkers')

    netz_biofuel_2 = netz_biofuel_1.append([biofuel_bunkers]).reset_index(drop = True)

    netz_biofuel_2.loc[netz_biofuel_2['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_biofuel_2.loc[netz_biofuel_2['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_biofuel_2.loc[netz_biofuel_2['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    netz_biofuel_2.loc[netz_biofuel_2['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    netz_biofuel_2.loc[netz_biofuel_2['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    netz_biofuel_2 = netz_biofuel_2[netz_biofuel_2['item_code_new'].isin(fuel_final_bunk)].reset_index(drop = True)

    netz_biofuel_2['item_code_new'] = pd.Categorical(
        netz_biofuel_2['item_code_new'], 
        categories = fuel_final_bunk, 
        ordered = True)

    netz_biofuel_2 = netz_biofuel_2.sort_values('item_code_new').reset_index(drop = True)

    netz_biofuel_2_rows = netz_biofuel_2.shape[0]
    netz_biofuel_2_cols = netz_biofuel_2.shape[1]

    # liquid and solid renewables

    netz_renew_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(fuel_vector_1)) &
                                        (EGEDA_years_netzero['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', 
                                                                                '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                '16_8_other_liquid_biofuels']))]\
                                           .copy().groupby(['item_code_new']).sum().assign(fuel_code = 'Liquid and solid renewables').reset_index()\
                                           [['fuel_code', 'item_code_new'] + col_chart_years] 

    renew_bunkers = netz_renew_1[netz_renew_1['item_code_new'].isin(['4_international_marine_bunkers',
                                                                         '5_international_aviation_bunkers'])]\
                                                                             .groupby(['fuel_code']).sum().assign(fuel_code = 'Liquid and solid renewables',
                                                                                                                  item_code_new = 'Bunkers')

    netz_renew_2 = netz_renew_1.append([renew_bunkers]).reset_index(drop = True)

    netz_renew_2.loc[netz_renew_2['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_renew_2.loc[netz_renew_2['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_renew_2.loc[netz_renew_2['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    netz_renew_2.loc[netz_renew_2['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    netz_renew_2.loc[netz_renew_2['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    netz_renew_2 = netz_renew_2[netz_renew_2['item_code_new'].isin(fuel_final_bunk)].reset_index(drop = True)

    netz_renew_2['item_code_new'] = pd.Categorical(
        netz_renew_2['item_code_new'], 
        categories = fuel_final_bunk, 
        ordered = True)

    netz_renew_2 = netz_renew_2.sort_values('item_code_new').reset_index(drop = True)

    netz_renew_2_rows = netz_renew_2.shape[0]
    netz_renew_2_cols = netz_renew_2.shape[1]

    ###########################################################################################

    # Fuel consummption data frame builds
    # REFERENCE

    # Industry
    # Transport
    # Buildings
    # Agriculture
    # Non-specified
    # Non-energy
    # Own-use
    # Power (including heat)
    # Total

    # Coal

    ref_coal_ind = ref_ind_2[ref_ind_2['fuel_code'] == 'Coal']
    ref_coal_bld = ref_bld_2[ref_bld_2['fuel_code'] == 'Coal']
    ref_coal_ag = ref_ag_1[ref_ag_1['fuel_code'] == 'Coal']

    ref_coal_trn = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                         (EGEDA_years_reference['item_code_new'].isin(['15_transport_sector'])) &
                                         (EGEDA_years_reference['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                            '1_x_coal_thermal', '2_coal_products']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Coal')

    ref_coal_trn = ref_coal_trn[['fuel_code', 'item_code_new'] + list(ref_coal_trn.loc[:, '2000':'2050'])]

    ref_coal_ne = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(['17_nonenergy_use'])) &
                                        (EGEDA_years_reference['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                            '1_x_coal_thermal', '2_coal_products']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Coal')

    ref_coal_ne = ref_coal_ne[['fuel_code', 'item_code_new'] + list(ref_coal_ne.loc[:, '2000':'2050'])]

    ref_coal_ns = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                        (EGEDA_years_reference['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                            '1_x_coal_thermal', '2_coal_products']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Coal')

    ref_coal_ns = ref_coal_ns[['fuel_code', 'item_code_new'] + list(ref_coal_ns.loc[:, '2000':'2050'])]

    ref_coal_own = ref_ownuse_1[ref_ownuse_1['FUEL'] == 'Coal']
    ref_coal_own = ref_coal_own.rename(columns = {'FUEL': 'fuel_code', 'Sector': 'item_code_new'})

    ref_coal_pow = ref_pow_use_2[ref_pow_use_2['FUEL'].isin(['Coal', 'Lignite'])].copy().groupby(['Transformation']).sum()\
                        .reset_index(drop = True).assign(fuel_code = 'Coal', item_code_new = 'Power')

    ref_coal_pow = ref_coal_pow.rename(columns = {'FUEL': 'fuel_code', 'Transformation': 'item_code_new'})

    # Hydrogen
    ref_coal_hyd = ref_hyd_use_1[ref_hyd_use_1['FUEL'] == 'Coal'].copy()\
        .rename(columns = {'FUEL': 'fuel_code', 'TECHNOLOGY': 'item_code_new'})\
            .reset_index(drop = True)

    ref_coal_hyd.loc[ref_coal_hyd['item_code_new'] == 'Input fuel', 'item_code_new'] = 'Hydrogen'

    if ref_coal_hyd.empty:
        hyd_series = ['Coal', 'Hydrogen'] + [0] * 33
        hyd_grab = pd.Series(hyd_series, index = ref_coal_hyd.columns)
        ref_coal_hyd = ref_coal_hyd.append(hyd_grab, ignore_index = True)

    else:
        pass

    ref_coalcons_1 = ref_coal_ind.append([ref_coal_bld, ref_coal_ag, ref_coal_trn, ref_coal_ne, 
                                          ref_coal_ns, ref_coal_own, ref_coal_pow, ref_coal_hyd])\
                                              .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_coalcons_1.loc[ref_coalcons_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_coalcons_1.loc[ref_coalcons_1['item_code_new'] == '16_x_buildings', 'item_code_new'] = 'Buildings'
    ref_coalcons_1.loc[ref_coalcons_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    ref_coalcons_1.loc[ref_coalcons_1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    ref_coalcons_1.loc[ref_coalcons_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    ref_coalcons_1.loc['Total'] = ref_coalcons_1.sum(numeric_only = True)

    ref_coalcons_1.loc['Total', 'fuel_code'] = 'Coal'
    ref_coalcons_1.loc['Total', 'item_code_new'] = 'Total'

    ref_coalcons_1 = ref_coalcons_1.copy().reset_index(drop = True)

    ref_coalcons_1_rows = ref_coalcons_1.shape[0]
    ref_coalcons_1_cols = ref_coalcons_1.shape[1]

    # Coal consumption by type

    # Grabbing TPES as a proxy for demand

    ref_coaltpes_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                           (EGEDA_years_reference['item_code_new'].isin(['7_total_primary_energy_supply'])) &
                                           (EGEDA_years_reference['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                               '1_x_coal_thermal']))].copy().reset_index(drop = True)
    
    met_coal = ref_coaltpes_1[ref_coaltpes_1['fuel_code'].isin(['1_1_coking_coal', '2_coal_products'])].copy()\
         .groupby(['item_code_new']).sum().assign(fuel_code = 'Metallurgical coal', economy = economy).reset_index()

    ref_coaltpes_1 = ref_coaltpes_1.append(met_coal).reset_index(drop = True)

    ref_coaltpes_1.loc[ref_coaltpes_1['fuel_code'] == '1_x_coal_thermal', 'fuel_code'] = 'Thermal coal'
    ref_coaltpes_1.loc[ref_coaltpes_1['fuel_code'] == '1_5_lignite', 'fuel_code'] = 'Lignite'
    ref_coaltpes_1.loc[ref_coaltpes_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'TPES'    

    ref_coaltpes_2 = ref_coaltpes_1[['economy', 'fuel_code', 'item_code_new'] + list(ref_coaltpes_1.loc[:,'2000':'2050'])]

    ref_coaltpes_2 = ref_coaltpes_2[ref_coaltpes_2['fuel_code'].isin(['Thermal coal', 'Lignite', 'Metallurgical coal'])]\
        .copy().reset_index(drop = True)

    ref_coaltpes_2_rows = ref_coaltpes_2.shape[0]
    ref_coaltpes_2_cols = ref_coaltpes_2.shape[1]

    # Natural gas

    ref_gas_ind = ref_ind_2[ref_ind_2['fuel_code'] == 'Gas']
    ref_gas_bld = ref_bld_2[ref_bld_2['fuel_code'] == 'Gas']
    ref_gas_ag = ref_ag_1[ref_ag_1['fuel_code'] == 'Gas']
    ref_gas_trn = ref_trn_1[ref_trn_1['fuel_code'] == 'Gas']

    ref_gas_ne = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(['17_nonenergy_use'])) &
                                        (EGEDA_years_reference['fuel_code'].isin(['8_gas']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Gas')

    ref_gas_ne = ref_gas_ne[['fuel_code', 'item_code_new'] + list(ref_gas_ne.loc[:, '2000':'2050'])]

    ref_gas_ns = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                        (EGEDA_years_reference['fuel_code'].isin(['8_gas']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Gas')

    ref_gas_ns = ref_gas_ns[['fuel_code', 'item_code_new'] + list(ref_gas_ns.loc[:, '2000':'2050'])]

    ref_gas_own = ref_ownuse_1[ref_ownuse_1['FUEL'] == 'Gas']
    ref_gas_own = ref_gas_own.rename(columns = {'FUEL': 'fuel_code', 'Sector': 'item_code_new'})

    ref_gas_pow = ref_pow_use_2[ref_pow_use_2['FUEL'] == 'Gas']
    ref_gas_pow = ref_gas_pow.rename(columns = {'FUEL': 'fuel_code', 'Transformation': 'item_code_new'})

    # Hydrogen
    ref_gas_hyd = ref_hyd_use_1[ref_hyd_use_1['FUEL'] == 'Gas'].copy()\
        .rename(columns = {'FUEL': 'fuel_code', 'TECHNOLOGY': 'item_code_new'})\
            .reset_index(drop = True)

    ref_gas_hyd.loc[ref_gas_hyd['item_code_new'] == 'Input fuel', 'item_code_new'] = 'Hydrogen'

    if ref_gas_hyd.empty:
        hyd_series = ['Gas', 'Hydrogen'] + [0] * 33
        hyd_grab = pd.Series(hyd_series, index = ref_gas_hyd.columns)
        ref_gas_hyd = ref_gas_hyd.append(hyd_grab, ignore_index = True)

    else:
        pass

    ref_gascons_1 = ref_gas_ind.append([ref_gas_bld, ref_gas_ag, ref_gas_trn, ref_gas_ne, 
                                          ref_gas_ns, ref_gas_own, ref_gas_pow, ref_gas_hyd])\
                                              .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_gascons_1.loc[ref_gascons_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_gascons_1.loc[ref_gascons_1['item_code_new'] == '16_x_buildings', 'item_code_new'] = 'Buildings'
    ref_gascons_1.loc[ref_gascons_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    ref_gascons_1.loc[ref_gascons_1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    ref_gascons_1.loc[ref_gascons_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'
    ref_gascons_1.loc[ref_gascons_1['item_code_new'] == 'Input fuel', 'item_code_new'] = 'Power'

    ref_gascons_1.loc['Total'] = ref_gascons_1.sum(numeric_only = True)

    ref_gascons_1.loc['Total', 'fuel_code'] = 'Gas'
    ref_gascons_1.loc['Total', 'item_code_new'] = 'Total'

    ref_gascons_1 = ref_gascons_1.copy().reset_index(drop = True)

    ref_gascons_1_rows = ref_gascons_1.shape[0]
    ref_gascons_1_cols = ref_gascons_1.shape[1]

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

    # Petroleum products

    ref_petprod_ind = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                            (EGEDA_years_reference['item_code_new'].isin(['14_industry_sector'])) &
                                            (EGEDA_years_reference['fuel_code'].isin(['7_petroleum_products']))]\
                                                .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_petprod_ind = ref_petprod_ind[['fuel_code', 'item_code_new'] + list(ref_petprod_ind.loc[:, '2000':'2050'])]
    ref_petprod_ind.loc[ref_petprod_ind['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'

    ref_petprod_bld = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                            (EGEDA_years_reference['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                            (EGEDA_years_reference['fuel_code'].isin(['7_petroleum_products']))].copy().replace(np.nan, 0).groupby(['fuel_code'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Petroleum products', item_code_new = 'Buildings')

    ref_petprod_bld = ref_petprod_bld[['fuel_code', 'item_code_new'] + list(ref_petprod_bld.loc[:, '2000':'2050'])]

    ref_petprod_ag = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                           (EGEDA_years_reference['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])) &
                                           (EGEDA_years_reference['fuel_code'].isin(['7_petroleum_products']))].copy().replace(np.nan, 0).groupby(['fuel_code'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Petroleum products', item_code_new = 'Agriculture')

    ref_petprod_ag = ref_petprod_ag[['fuel_code', 'item_code_new'] + list(ref_petprod_ag.loc[:, '2000':'2050'])]

    ref_petprod_trn = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                            (EGEDA_years_reference['item_code_new'].isin(['15_transport_sector'])) &
                                            (EGEDA_years_reference['fuel_code'].isin(['7_petroleum_products']))]\
                                                .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_petprod_trn = ref_petprod_trn[['fuel_code', 'item_code_new'] + list(ref_petprod_trn.loc[:, '2000':'2050'])]
    ref_petprod_trn.loc[ref_petprod_trn['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'

    ref_petprod_ne = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                           (EGEDA_years_reference['item_code_new'].isin(['17_nonenergy_use'])) &
                                           (EGEDA_years_reference['fuel_code'].isin(['7_petroleum_products']))]\
                                               .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_petprod_ne = ref_petprod_ne[['fuel_code', 'item_code_new'] + list(ref_petprod_ne.loc[:, '2000':'2050'])]
    ref_petprod_ne.loc[ref_petprod_ne['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'

    ref_petprod_ns = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                           (EGEDA_years_reference['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                           (EGEDA_years_reference['fuel_code'].isin(['7_petroleum_products']))]\
                                               .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_petprod_ns = ref_petprod_ns[['fuel_code', 'item_code_new'] + list(ref_petprod_ns.loc[:, '2000':'2050'])]
    ref_petprod_ns.loc[ref_petprod_ns['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'

    # Own-use
    ref_petprod_own = ref_trans_df1[(ref_trans_df1['economy'] == economy) & 
                                    (ref_trans_df1['Sector'] == 'OWN') &
                                    (ref_trans_df1['FUEL'].isin(['7_1_motor_gasoline', '7_2_aviation_gasoline', '7_3_naphtha',
                                                                 '7_x_jet_fuel', '7_6_kerosene', '7_7_gas_diesel_oil', '7_8_fuel_oil',
                                                                 '7_9_lpg', '7_10_refinery_gas_not_liquefied', '7_11_ethane',
                                                                 '7_x_other_petroleum_products']))]\
                                                                     .copy().reset_index(drop = True)

    ref_petprod_own = ref_petprod_own.groupby(['economy']).sum().copy().reset_index(drop = True)\
                        .assign(fuel_code = 'Petroleum products', item_code_new = '10_losses_and_own_use')

    #################################################################################
    hist_ownoil = EGEDA_hist_own_oil[(EGEDA_hist_own_oil['economy'] == economy) &
                                     (EGEDA_hist_own_oil['FUEL'] == 'Petroleum products')].copy().\
                                        iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_hist_own_oil.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_petprod_own = hist_ownoil.merge(ref_petprod_own[['fuel_code', 'item_code_new'] + list(ref_petprod_own.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_petprod_own = ref_petprod_own[['fuel_code', 'item_code_new'] + list(ref_petprod_own.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    # Power
    ref_petprod_power = ref_power_df1[(ref_power_df1['economy'] == economy) &
                                    (ref_power_df1['FUEL'].isin(['7_3_naphtha', '7_7_gas_diesel_oil', '7_8_fuel_oil',
                                                                 '7_9_lpg', '7_10_refinery_gas_not_liquefied',
                                                                 '7_x_other_petroleum_products', '7_16_petroleum_coke']))]\
                                                                     .copy().reset_index(drop = True)

    ref_petprod_power = ref_petprod_power.groupby(['economy']).sum().copy().reset_index(drop = True)\
                            .assign(fuel_code = 'Petroleum products', item_code_new = 'Power')

    #################################################################################
    hist_poweroil = EGEDA_histpower_oil[(EGEDA_histpower_oil['economy'] == economy) &
                                        (EGEDA_histpower_oil['FUEL'] == 'Petroleum products')].copy()\
                                            .iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_histpower_oil.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_petprod_power = hist_poweroil.merge(ref_petprod_power[['fuel_code', 'item_code_new'] + list(ref_petprod_power.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_petprod_power = ref_petprod_power[['fuel_code', 'item_code_new'] + list(ref_petprod_power.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    ref_petprodcons_1 = ref_petprod_ind.append([ref_petprod_bld, ref_petprod_ag, ref_petprod_trn, ref_petprod_ne, 
                                                ref_petprod_ns, ref_petprod_own, ref_petprod_power])\
                                                    .copy().reset_index(drop = True)

    ref_petprodcons_1.loc[ref_petprodcons_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own-use and losses'
    ref_petprodcons_1.loc[ref_petprodcons_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_petprodcons_1.loc[ref_petprodcons_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    ref_petprodcons_1.loc[ref_petprodcons_1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    ref_petprodcons_1.loc[ref_petprodcons_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    ref_petprodcons_1.loc['Total'] = ref_petprodcons_1.sum(numeric_only = True)

    ref_petprodcons_1.loc['Total', 'fuel_code'] = 'Petroleum products'
    ref_petprodcons_1.loc['Total', 'item_code_new'] = 'Total'

    ref_petprodcons_1 = ref_petprodcons_1.copy().reset_index(drop = True)
    
    ref_petprodcons_1_rows = ref_petprodcons_1.shape[0]
    ref_petprodcons_1_cols = ref_petprodcons_1.shape[1]

    # Liquid and solid renewables

    ref_renew_ind = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                            (EGEDA_years_reference['item_code_new'].isin(['14_industry_sector'])) &
                                            (EGEDA_years_reference['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', 
                                                                                      '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                      '16_8_other_liquid_biofuels']))]\
                                                .copy().replace(np.nan, 0).groupby(['item_code_new']).sum().reset_index(drop = True)\
                                                    .assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Industry')

    ref_renew_ind = ref_renew_ind[['fuel_code', 'item_code_new'] + list(ref_renew_ind.loc[:, '2000':'2050'])]

    ref_renew_bld = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                            (EGEDA_years_reference['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                            (EGEDA_years_reference['fuel_code'].isin(['16_1_biogas', '16_3_municipal_solid_waste_renewable', 
                                                                                      '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                      '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Buildings')

    ref_renew_bld = ref_renew_bld[['fuel_code', 'item_code_new'] + list(ref_renew_bld.loc[:, '2000':'2050'])]

    ref_renew_bldtrad = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                              (EGEDA_years_reference['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                              (EGEDA_years_reference['fuel_code'].isin(['15_solid_biomass']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                  .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Buildings (biomass)')

    ref_renew_bldtrad = ref_renew_bldtrad[['fuel_code', 'item_code_new'] + list(ref_renew_bldtrad.loc[:, '2000':'2050'])]

    # Independent build for ref_renewcons_2 #######################################################################################################################
    ref_renew_bldall = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                             (EGEDA_years_reference['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                             (EGEDA_years_reference['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', 
                                                                                       '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                       '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                 .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Buildings')

    ref_renew_bldall = ref_renew_bldall[['fuel_code', 'item_code_new'] + list(ref_renew_bldall.loc[:, '2000':'2050'])]
    ###############################################################################################################################################################

    ref_renew_ag = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                           (EGEDA_years_reference['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])) &
                                           (EGEDA_years_reference['fuel_code'].isin(['15_solid_biomass', '16_1_biogas','16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                     '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                     '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Agriculture')

    ref_renew_ag = ref_renew_ag[['fuel_code', 'item_code_new'] + list(ref_renew_ag.loc[:, '2000':'2050'])]

    ref_renew_trn = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                            (EGEDA_years_reference['item_code_new'].isin(['15_transport_sector'])) &
                                            (EGEDA_years_reference['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', 
                                                                                      '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                      '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Transport')

    ref_renew_trn = ref_renew_trn[['fuel_code', 'item_code_new'] + list(ref_renew_trn.loc[:, '2000':'2050'])]

    ref_renew_ne = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                           (EGEDA_years_reference['item_code_new'].isin(['17_nonenergy_use'])) &
                                           (EGEDA_years_reference['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', 
                                                                                     '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                     '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Non-energy')

    ref_renew_ne = ref_renew_ne[['fuel_code', 'item_code_new'] + list(ref_renew_ne.loc[:, '2000':'2050'])]
    
    ref_renew_ns = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                           (EGEDA_years_reference['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                           (EGEDA_years_reference['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', 
                                                                                     '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                     '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Non-specified')

    ref_renew_ns = ref_renew_ns[['fuel_code', 'item_code_new'] + list(ref_renew_ns.loc[:, '2000':'2050'])]

    # Own-use
    ref_renew_own = ref_trans_df1[(ref_trans_df1['economy'] == economy) & 
                                    (ref_trans_df1['Sector'] == 'OWN') &
                                    (ref_trans_df1['FUEL'].isin(['15_1_fuelwood_and_woodwaste', '15_2_bagasse', '15_3_charcoal', '15_5_other_biomass', 
                                                                 '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                 '16_8_other_liquid_biofuels']))]\
                                                                     .copy().reset_index(drop = True)
                                                                     
    ref_renew_own = ref_renew_own.groupby(['economy']).sum().copy().reset_index(drop = True)\
                        .assign(fuel_code = 'Liquid and solid renewables', item_code_new = '10_losses_and_own_use')

    #################################################################################
    hist_ownrenew = EGEDA_hist_own_renew[(EGEDA_hist_own_renew['economy'] == economy) &
                                         (EGEDA_hist_own_renew['FUEL'] == 'Liquid and solid renewables')].copy().reset_index(drop = True)\
                                             .groupby(['economy']).sum().reset_index(drop = True)\
                                                 .assign(fuel_code = 'Liquid and solid renewables', item_code_new = '10_losses_and_own_use')\
                                                     [['fuel_code', 'item_code_new'] + list(EGEDA_hist_own_renew.loc[:, '2000':'2018'])]

    ref_renew_own = hist_ownrenew.merge(ref_renew_own[['fuel_code', 'item_code_new'] + list(ref_renew_own.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_renew_own = ref_renew_own[['fuel_code', 'item_code_new'] + list(ref_renew_own.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    # Power
    ref_renew_power = ref_power_df1[(ref_power_df1['economy'] == economy) &
                                    (ref_power_df1['FUEL'].isin(['15_4_black_liquor', '15_5_other_biomass', '16_3_municipal_solid_waste_renewable']))]\
                                                                     .copy().reset_index(drop = True)

    ref_renew_power = ref_renew_power.groupby(['economy']).sum().copy().reset_index(drop = True)\
                            .assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Power')

    #################################################################################
    hist_powerrenew = EGEDA_histpower_renew[(EGEDA_histpower_renew['economy'] == economy) &
                                        (EGEDA_histpower_renew['FUEL'] == 'Liquid and solid renewables')].copy()\
                                            .iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_histpower_renew.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_renew_power = hist_powerrenew.merge(ref_renew_power[['fuel_code', 'item_code_new'] + list(ref_renew_power.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_renew_power = ref_renew_power[['fuel_code', 'item_code_new'] + list(ref_renew_power.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    ref_renewcons_1 = ref_renew_ind.append([ref_renew_bld, ref_renew_bldtrad, ref_renew_ag, ref_renew_trn, ref_renew_ne, 
                                                ref_renew_ns, ref_renew_own, ref_renew_power])\
                                                    .copy().reset_index(drop = True)

    ref_renewcons_1.loc[ref_renewcons_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own-use and losses'

    ref_renewcons_1.loc['Total'] = ref_renewcons_1.sum(numeric_only = True)

    ref_renewcons_1.loc['Total', 'fuel_code'] = 'Liquid and solid renewables'
    ref_renewcons_1.loc['Total', 'item_code_new'] = 'Total'

    ref_renewcons_1 = ref_renewcons_1.copy().reset_index(drop = True)
    
    ref_renewcons_1_rows = ref_renewcons_1.shape[0]
    ref_renewcons_1_cols = ref_renewcons_1.shape[1]

    # Alternative that just has buildings as one category

    ref_renewcons_2 = ref_renew_ind.append([ref_renew_bldall, ref_renew_ag, ref_renew_trn, ref_renew_ne, 
                                                ref_renew_ns, ref_renew_own, ref_renew_power])\
                                                    .copy().reset_index(drop = True)

    ref_renewcons_2.loc[ref_renewcons_2['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own-use and losses'

    ref_renewcons_2.loc['Total'] = ref_renewcons_2.sum(numeric_only = True)

    ref_renewcons_2.loc['Total', 'fuel_code'] = 'Liquid and solid renewables'
    ref_renewcons_2.loc['Total', 'item_code_new'] = 'Total'

    ref_renewcons_2 = ref_renewcons_2.copy().reset_index(drop = True)
    
    ref_renewcons_2_rows = ref_renewcons_2.shape[0]
    ref_renewcons_2_cols = ref_renewcons_2.shape[1]

    ##########
    # CARBON NEUTRALITY

    # Industry
    # Transport
    # Buildings
    # Agriculture
    # Non-specified
    # Non-energy
    # Own-use
    # Power (including heat)
    # Refining for petroleum products
    # Total

    # Coal

    netz_coal_ind = netz_ind_2[netz_ind_2['fuel_code'] == 'Coal']
    netz_coal_bld = netz_bld_2[netz_bld_2['fuel_code'] == 'Coal']
    netz_coal_ag = netz_ag_1[netz_ag_1['fuel_code'] == 'Coal']

    netz_coal_trn = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                         (EGEDA_years_netzero['item_code_new'].isin(['15_transport_sector'])) &
                                         (EGEDA_years_netzero['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                            '1_x_coal_thermal', '2_coal_products']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Coal')

    netz_coal_trn = netz_coal_trn[['fuel_code', 'item_code_new'] + list(netz_coal_trn.loc[:, '2000':'2050'])]

    netz_coal_ne = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(['17_nonenergy_use'])) &
                                        (EGEDA_years_netzero['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                            '1_x_coal_thermal', '2_coal_products']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Coal')

    netz_coal_ne = netz_coal_ne[['fuel_code', 'item_code_new'] + list(netz_coal_ne.loc[:, '2000':'2050'])]

    netz_coal_ns = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                        (EGEDA_years_netzero['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                            '1_x_coal_thermal', '2_coal_products']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Coal')

    netz_coal_ns = netz_coal_ns[['fuel_code', 'item_code_new'] + list(netz_coal_ns.loc[:, '2000':'2050'])]

    netz_coal_own = netz_ownuse_1[netz_ownuse_1['FUEL'] == 'Coal']
    netz_coal_own = netz_coal_own.rename(columns = {'FUEL': 'fuel_code', 'Sector': 'item_code_new'})

    netz_coal_pow = netz_pow_use_2[netz_pow_use_2['FUEL'].isin(['Coal', 'Lignite'])].copy().groupby(['Transformation']).sum()\
                        .reset_index(drop = True).assign(fuel_code = 'Coal', item_code_new = 'Power')

    netz_coal_pow = netz_coal_pow.rename(columns = {'FUEL': 'fuel_code', 'Transformation': 'item_code_new'})

    # Hydrogen
    netz_coal_hyd = netz_hyd_use_1[netz_hyd_use_1['FUEL'] == 'Coal'].copy()\
        .rename(columns = {'FUEL': 'fuel_code', 'TECHNOLOGY': 'item_code_new'})\
            .reset_index(drop = True)

    netz_coal_hyd.loc[netz_coal_hyd['item_code_new'] == 'Input fuel', 'item_code_new'] = 'Hydrogen'

    if netz_coal_hyd.empty:
        hyd_series = ['Coal', 'Hydrogen'] + [0] * 33
        hyd_grab = pd.Series(hyd_series, index = netz_coal_hyd.columns)
        netz_coal_hyd = netz_coal_hyd.append(hyd_grab, ignore_index = True)

    else:
        pass

    netz_coalcons_1 = netz_coal_ind.append([netz_coal_bld, netz_coal_ag, netz_coal_trn, netz_coal_ne, 
                                          netz_coal_ns, netz_coal_own, netz_coal_pow, netz_coal_hyd])\
                                              .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_coalcons_1.loc[netz_coalcons_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_coalcons_1.loc[netz_coalcons_1['item_code_new'] == '16_x_buildings', 'item_code_new'] = 'Buildings'
    netz_coalcons_1.loc[netz_coalcons_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    netz_coalcons_1.loc[netz_coalcons_1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    netz_coalcons_1.loc[netz_coalcons_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    netz_coalcons_1.loc['Total'] = netz_coalcons_1.sum(numeric_only = True)

    netz_coalcons_1.loc['Total', 'fuel_code'] = 'Coal'
    netz_coalcons_1.loc['Total', 'item_code_new'] = 'Total'

    netz_coalcons_1 = netz_coalcons_1.copy().reset_index(drop = True)

    netz_coalcons_1_rows = netz_coalcons_1.shape[0]
    netz_coalcons_1_cols = netz_coalcons_1.shape[1]

    # Grabbing TPES as a proxy for demand

    netz_coaltpes_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                           (EGEDA_years_netzero['item_code_new'].isin(['7_total_primary_energy_supply'])) &
                                           (EGEDA_years_netzero['fuel_code'].isin(['1_1_coking_coal', '1_5_lignite',\
                                               '1_x_coal_thermal']))].copy().reset_index(drop = True)
    
    met_coal = netz_coaltpes_1[netz_coaltpes_1['fuel_code'].isin(['1_1_coking_coal', '2_coal_products'])].copy()\
         .groupby(['item_code_new']).sum().assign(fuel_code = 'Metallurgical coal', economy = economy).reset_index()

    netz_coaltpes_1 = netz_coaltpes_1.append(met_coal).reset_index(drop = True)

    netz_coaltpes_1.loc[netz_coaltpes_1['fuel_code'] == '1_x_coal_thermal', 'fuel_code'] = 'Thermal coal'
    netz_coaltpes_1.loc[netz_coaltpes_1['fuel_code'] == '1_5_lignite', 'fuel_code'] = 'Lignite'
    netz_coaltpes_1.loc[netz_coaltpes_1['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'TPES'

    netz_coaltpes_2 = netz_coaltpes_1[['economy', 'fuel_code', 'item_code_new'] + list(netz_coaltpes_1.loc[:,'2000':'2050'])]

    netz_coaltpes_2 = netz_coaltpes_2[netz_coaltpes_2['fuel_code'].isin(['Thermal coal', 'Lignite', 'Metallurgical coal'])]\
        .copy().reset_index(drop = True)

    netz_coaltpes_2_rows = netz_coaltpes_2.shape[0]
    netz_coaltpes_2_cols = netz_coaltpes_2.shape[1]

    # Natural gas

    netz_gas_ind = netz_ind_2[netz_ind_2['fuel_code'] == 'Gas']
    netz_gas_bld = netz_bld_2[netz_bld_2['fuel_code'] == 'Gas']
    netz_gas_ag = netz_ag_1[netz_ag_1['fuel_code'] == 'Gas']
    netz_gas_trn = netz_trn_1[netz_trn_1['fuel_code'] == 'Gas']

    netz_gas_ne = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(['17_nonenergy_use'])) &
                                        (EGEDA_years_netzero['fuel_code'].isin(['8_gas']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Gas')

    netz_gas_ne = netz_gas_ne[['fuel_code', 'item_code_new'] + list(netz_gas_ne.loc[:, '2000':'2050'])]

    netz_gas_ns = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                        (EGEDA_years_netzero['fuel_code'].isin(['8_gas']))].copy().groupby(['item_code_new'])\
                                                .sum().reset_index().assign(fuel_code = 'Gas')

    netz_gas_ns = netz_gas_ns[['fuel_code', 'item_code_new'] + list(netz_gas_ns.loc[:, '2000':'2050'])]

    netz_gas_own = netz_ownuse_1[netz_ownuse_1['FUEL'] == 'Gas']
    netz_gas_own = netz_gas_own.rename(columns = {'FUEL': 'fuel_code', 'Sector': 'item_code_new'})

    netz_gas_pow = netz_pow_use_2[netz_pow_use_2['FUEL'] == 'Gas']
    netz_gas_pow = netz_gas_pow.rename(columns = {'FUEL': 'fuel_code', 'Transformation': 'item_code_new'})

    # Hydrogen
    netz_gas_hyd = netz_hyd_use_1[netz_hyd_use_1['FUEL'] == 'Gas'].copy()\
        .rename(columns = {'FUEL': 'fuel_code', 'TECHNOLOGY': 'item_code_new'})\
            .reset_index(drop = True)

    netz_gas_hyd.loc[netz_gas_hyd['item_code_new'] == 'Input fuel', 'item_code_new'] = 'Hydrogen'

    if netz_gas_hyd.empty:
        hyd_series = ['Gas', 'Hydrogen'] + [0] * 33
        hyd_grab = pd.Series(hyd_series, index = netz_gas_hyd.columns)
        netz_gas_hyd = netz_gas_hyd.append(hyd_grab, ignore_index = True)

    else:
        pass

    netz_gascons_1 = netz_gas_ind.append([netz_gas_bld, netz_gas_ag, netz_gas_trn, netz_gas_ne, 
                                          netz_gas_ns, netz_gas_own, netz_gas_pow, netz_gas_hyd])\
                                              .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_gascons_1.loc[netz_gascons_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_gascons_1.loc[netz_gascons_1['item_code_new'] == '16_x_buildings', 'item_code_new'] = 'Buildings'
    netz_gascons_1.loc[netz_gascons_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    netz_gascons_1.loc[netz_gascons_1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    netz_gascons_1.loc[netz_gascons_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'
    netz_gascons_1.loc[netz_gascons_1['item_code_new'] == 'Input fuel', 'item_code_new'] = 'Power'

    netz_gascons_1.loc['Total'] = netz_gascons_1.sum(numeric_only = True)

    netz_gascons_1.loc['Total', 'fuel_code'] = 'Gas'
    netz_gascons_1.loc['Total', 'item_code_new'] = 'Total'

    netz_gascons_1 = netz_gascons_1.copy().reset_index(drop = True)

    netz_gascons_1_rows = netz_gascons_1.shape[0]
    netz_gascons_1_cols = netz_gascons_1.shape[1]

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

    # Petroleum products

    netz_petprod_ind = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                            (EGEDA_years_netzero['item_code_new'].isin(['14_industry_sector'])) &
                                            (EGEDA_years_netzero['fuel_code'].isin(['7_petroleum_products']))]\
                                                .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_petprod_ind = netz_petprod_ind[['fuel_code', 'item_code_new'] + list(netz_petprod_ind.loc[:, '2000':'2050'])]
    netz_petprod_ind.loc[netz_petprod_ind['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'

    netz_petprod_bld = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                            (EGEDA_years_netzero['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                            (EGEDA_years_netzero['fuel_code'].isin(['7_petroleum_products']))].copy().replace(np.nan, 0).groupby(['fuel_code'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Petroleum products', item_code_new = 'Buildings')

    netz_petprod_bld = netz_petprod_bld[['fuel_code', 'item_code_new'] + list(netz_petprod_bld.loc[:, '2000':'2050'])]

    netz_petprod_ag = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                           (EGEDA_years_netzero['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])) &
                                           (EGEDA_years_netzero['fuel_code'].isin(['7_petroleum_products']))].copy().replace(np.nan, 0).groupby(['fuel_code'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Petroleum products', item_code_new = 'Agriculture')

    netz_petprod_ag = netz_petprod_ag[['fuel_code', 'item_code_new'] + list(netz_petprod_ag.loc[:, '2000':'2050'])]

    netz_petprod_trn = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                            (EGEDA_years_netzero['item_code_new'].isin(['15_transport_sector'])) &
                                            (EGEDA_years_netzero['fuel_code'].isin(['7_petroleum_products']))]\
                                                .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_petprod_trn = netz_petprod_trn[['fuel_code', 'item_code_new'] + list(netz_petprod_trn.loc[:, '2000':'2050'])]
    netz_petprod_trn.loc[netz_petprod_trn['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'

    netz_petprod_ne = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                           (EGEDA_years_netzero['item_code_new'].isin(['17_nonenergy_use'])) &
                                           (EGEDA_years_netzero['fuel_code'].isin(['7_petroleum_products']))]\
                                               .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_petprod_ne = netz_petprod_ne[['fuel_code', 'item_code_new'] + list(netz_petprod_ne.loc[:, '2000':'2050'])]
    netz_petprod_ne.loc[netz_petprod_ne['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'

    netz_petprod_ns = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                           (EGEDA_years_netzero['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                           (EGEDA_years_netzero['fuel_code'].isin(['7_petroleum_products']))]\
                                               .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_petprod_ns = netz_petprod_ns[['fuel_code', 'item_code_new'] + list(netz_petprod_ns.loc[:, '2000':'2050'])]
    netz_petprod_ns.loc[netz_petprod_ns['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'

    # Own-use
    netz_petprod_own = netz_trans_df1[(netz_trans_df1['economy'] == economy) & 
                                    (netz_trans_df1['Sector'] == 'OWN') &
                                    (netz_trans_df1['FUEL'].isin(['7_1_motor_gasoline', '7_2_aviation_gasoline', '7_3_naphtha',
                                                                 '7_x_jet_fuel', '7_6_kerosene', '7_7_gas_diesel_oil', '7_8_fuel_oil',
                                                                 '7_9_lpg', '7_10_refinery_gas_not_liquefied', '7_11_ethane',
                                                                 '7_x_other_petroleum_products']))]\
                                                                     .copy().reset_index(drop = True)

    netz_petprod_own = netz_petprod_own.groupby(['economy']).sum().copy().reset_index(drop = True)\
                        .assign(fuel_code = 'Petroleum products', item_code_new = '10_losses_and_own_use')

    #################################################################################
    hist_ownoil = EGEDA_hist_own_oil[(EGEDA_hist_own_oil['economy'] == economy) &
                                     (EGEDA_hist_own_oil['FUEL'] == 'Petroleum products')].copy().\
                                        iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_hist_own_oil.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_petprod_own = hist_ownoil.merge(netz_petprod_own[['fuel_code', 'item_code_new'] + list(netz_petprod_own.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_petprod_own = netz_petprod_own[['fuel_code', 'item_code_new'] + list(netz_petprod_own.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    # Power
    netz_petprod_power = netz_power_df1[(netz_power_df1['economy'] == economy) &
                                    (netz_power_df1['FUEL'].isin(['7_3_naphtha', '7_7_gas_diesel_oil', '7_8_fuel_oil',
                                                                 '7_9_lpg', '7_10_refinery_gas_not_liquefied',
                                                                 '7_x_other_petroleum_products', '7_16_petroleum_coke']))]\
                                                                     .copy().reset_index(drop = True)

    netz_petprod_power = netz_petprod_power.groupby(['economy']).sum().copy().reset_index(drop = True)\
                            .assign(fuel_code = 'Petroleum products', item_code_new = 'Power')

    #################################################################################
    hist_poweroil = EGEDA_histpower_oil[(EGEDA_histpower_oil['economy'] == economy) &
                                        (EGEDA_histpower_oil['FUEL'] == 'Petroleum products')].copy()\
                                            .iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_histpower_oil.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_petprod_power = hist_poweroil.merge(netz_petprod_power[['fuel_code', 'item_code_new'] + list(netz_petprod_power.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_petprod_power = netz_petprod_power[['fuel_code', 'item_code_new'] + list(netz_petprod_power.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    netz_petprodcons_1 = netz_petprod_ind.append([netz_petprod_bld, netz_petprod_ag, netz_petprod_trn, netz_petprod_ne, 
                                                netz_petprod_ns, netz_petprod_own, netz_petprod_power])\
                                                    .copy().reset_index(drop = True)

    netz_petprodcons_1.loc[netz_petprodcons_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own-use and losses'
    netz_petprodcons_1.loc[netz_petprodcons_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_petprodcons_1.loc[netz_petprodcons_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    netz_petprodcons_1.loc[netz_petprodcons_1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    netz_petprodcons_1.loc[netz_petprodcons_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    netz_petprodcons_1.loc['Total'] = netz_petprodcons_1.sum(numeric_only = True)

    netz_petprodcons_1.loc['Total', 'fuel_code'] = 'Petroleum products'
    netz_petprodcons_1.loc['Total', 'item_code_new'] = 'Total'

    netz_petprodcons_1 = netz_petprodcons_1.copy().reset_index(drop = True)
    
    netz_petprodcons_1_rows = netz_petprodcons_1.shape[0]
    netz_petprodcons_1_cols = netz_petprodcons_1.shape[1]

    # Liquid and solid renewables

    netz_renew_ind = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                            (EGEDA_years_netzero['item_code_new'].isin(['14_industry_sector'])) &
                                            (EGEDA_years_netzero['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))]\
                                                .copy().replace(np.nan, 0).groupby(['item_code_new']).sum().reset_index(drop = True)\
                                                    .assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Industry')

    netz_renew_ind = netz_renew_ind[['fuel_code', 'item_code_new'] + list(netz_renew_ind.loc[:, '2000':'2050'])]

    netz_renew_bld = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                            (EGEDA_years_netzero['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                            (EGEDA_years_netzero['fuel_code'].isin(['16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Buildings')

    netz_renew_bld = netz_renew_bld[['fuel_code', 'item_code_new'] + list(netz_renew_bld.loc[:, '2000':'2050'])]

    netz_renew_bldtrad = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                              (EGEDA_years_netzero['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                              (EGEDA_years_netzero['fuel_code'].isin(['15_solid_biomass']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                  .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Buildings (biomass)')

    netz_renew_bldtrad = netz_renew_bldtrad[['fuel_code', 'item_code_new'] + list(netz_renew_bldtrad.loc[:, '2000':'2050'])]

    # Independent build for netz_renewcons_2 #############################################################################################################################
    netz_renew_bldall = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                            (EGEDA_years_netzero['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                            (EGEDA_years_netzero['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                  .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Buildings')

    netz_renew_bldall = netz_renew_bldall[['fuel_code', 'item_code_new'] + list(netz_renew_bldall.loc[:, '2000':'2050'])]
    #####################################################################################################################################################################

    netz_renew_ag = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                           (EGEDA_years_netzero['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])) &
                                           (EGEDA_years_netzero['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Agriculture')

    netz_renew_ag = netz_renew_ag[['fuel_code', 'item_code_new'] + list(netz_renew_ag.loc[:, '2000':'2050'])]

    netz_renew_trn = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                            (EGEDA_years_netzero['item_code_new'].isin(['15_transport_sector'])) &
                                            (EGEDA_years_netzero['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Transport')

    netz_renew_trn = netz_renew_trn[['fuel_code', 'item_code_new'] + list(netz_renew_trn.loc[:, '2000':'2050'])]

    netz_renew_ne = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                           (EGEDA_years_netzero['item_code_new'].isin(['17_nonenergy_use'])) &
                                           (EGEDA_years_netzero['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Non-energy')

    netz_renew_ne = netz_renew_ne[['fuel_code', 'item_code_new'] + list(netz_renew_ne.loc[:, '2000':'2050'])]
    
    netz_renew_ns = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                           (EGEDA_years_netzero['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                           (EGEDA_years_netzero['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))].copy().replace(np.nan, 0).groupby(['economy'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Non-specified')

    netz_renew_ns = netz_renew_ns[['fuel_code', 'item_code_new'] + list(netz_renew_ns.loc[:, '2000':'2050'])]

    # Own-use
    netz_renew_own = netz_trans_df1[(netz_trans_df1['economy'] == economy) & 
                                    (netz_trans_df1['Sector'] == 'OWN') &
                                    (netz_trans_df1['FUEL'].isin(['15_1_fuelwood_and_woodwaste', '15_2_bagasse', '15_3_charcoal', '15_5_other_biomass', 
                                                                 '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                 '16_8_other_liquid_biofuels']))]\
                                                                     .copy().reset_index(drop = True)
                                                                     
    netz_renew_own = netz_renew_own.groupby(['economy']).sum().copy().reset_index(drop = True)\
                        .assign(fuel_code = 'Liquid and solid renewables', item_code_new = '10_losses_and_own_use')

    #################################################################################
    hist_ownrenew = EGEDA_hist_own_renew[(EGEDA_hist_own_renew['economy'] == economy) &
                                         (EGEDA_hist_own_renew['FUEL'] == 'Liquid and solid renewables')].copy().reset_index(drop = True)\
                                             .groupby(['economy']).sum().reset_index(drop = True)\
                                                 .assign(fuel_code = 'Liquid and solid renewables', item_code_new = '10_losses_and_own_use')\
                                                     [['fuel_code', 'item_code_new'] + list(EGEDA_hist_own_renew.loc[:, '2000':'2018'])]

    netz_renew_own = hist_ownrenew.merge(netz_renew_own[['fuel_code', 'item_code_new'] + list(netz_renew_own.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_renew_own = netz_renew_own[['fuel_code', 'item_code_new'] + list(netz_renew_own.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    # Power
    netz_renew_power = netz_power_df1[(netz_power_df1['economy'] == economy) &
                                    (netz_power_df1['FUEL'].isin(['15_4_black_liquor', '15_5_other_biomass', '16_3_municipal_solid_waste_renewable']))]\
                                                                     .copy().reset_index(drop = True)

    netz_renew_power = netz_renew_power.groupby(['economy']).sum().copy().reset_index(drop = True)\
                            .assign(fuel_code = 'Liquid and solid renewables', item_code_new = 'Power')

    #################################################################################
    hist_powerrenew = EGEDA_histpower_renew[(EGEDA_histpower_renew['economy'] == economy) &
                                        (EGEDA_histpower_renew['FUEL'] == 'Liquid and solid renewables')].copy()\
                                            .iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_histpower_renew.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_renew_power = hist_powerrenew.merge(netz_renew_power[['fuel_code', 'item_code_new'] + list(netz_renew_power.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_renew_power = netz_renew_power[['fuel_code', 'item_code_new'] + list(netz_renew_power.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    netz_renewcons_1 = netz_renew_ind.append([netz_renew_bld, netz_renew_bldtrad, netz_renew_ag, netz_renew_trn, netz_renew_ne, 
                                                netz_renew_ns, netz_renew_own, netz_renew_power])\
                                                    .copy().reset_index(drop = True)

    netz_renewcons_1.loc[netz_renewcons_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own-use and losses'

    netz_renewcons_1.loc['Total'] = netz_renewcons_1.sum(numeric_only = True)

    netz_renewcons_1.loc['Total', 'fuel_code'] = 'Liquid and solid renewables'
    netz_renewcons_1.loc['Total', 'item_code_new'] = 'Total'

    netz_renewcons_1 = netz_renewcons_1.copy().reset_index(drop = True)
    
    netz_renewcons_1_rows = netz_renewcons_1.shape[0]
    netz_renewcons_1_cols = netz_renewcons_1.shape[1]

    # Build for buildings in one category
    netz_renewcons_2 = netz_renew_ind.append([netz_renew_bldall, netz_renew_ag, netz_renew_trn, netz_renew_ne, 
                                                netz_renew_ns, netz_renew_own, netz_renew_power])\
                                                    .copy().reset_index(drop = True)

    netz_renewcons_2.loc[netz_renewcons_2['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own-use and losses'

    netz_renewcons_2.loc['Total'] = netz_renewcons_2.sum(numeric_only = True)

    netz_renewcons_2.loc['Total', 'fuel_code'] = 'Liquid and solid renewables'
    netz_renewcons_2.loc['Total', 'item_code_new'] = 'Total'

    netz_renewcons_2 = netz_renewcons_2.copy().reset_index(drop = True)
    
    netz_renewcons_2_rows = netz_renewcons_2.shape[0]
    netz_renewcons_2_cols = netz_renewcons_2.shape[1]

    ##########################

    ref_petprod_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(fuel_vector_ref)) &
                                        (EGEDA_years_reference['fuel_code'] == '7_petroleum_products')].copy()\
                                            [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                .reset_index(drop = True)
    
    ref_petprod_1['fuel_code'].replace({'7_petroleum_products': 'Petroleum products'}, inplace=True)

    petprod_bunkers = ref_petprod_1[ref_petprod_1['item_code_new'].isin(['4_international_marine_bunkers',
                                                                         '5_international_aviation_bunkers'])]\
                                                                             .groupby(['fuel_code']).sum().assign(fuel_code = 'Petroleum products',
                                                                                                                  item_code_new = 'Bunkers')
    
    ref_petprod_2 = ref_petprod_1.append([petprod_bunkers]).reset_index(drop = True)

    if 'Refining' in list(ref_crudecons_1['item_code_new']):
        refining_grab = ref_crudecons_1[ref_crudecons_1['item_code_new'] == 'Refining'][['fuel_code', 'item_code_new'] + col_chart_years].copy()

    else:
        refining_series = ['Crude oil & NGL', 'Refining'] + [0] * 7
        refining_grab = pd.Series(refining_series, index = ref_petprod_1.columns)

    ref_petprod_2 = ref_petprod_2.append(refining_grab, ignore_index = True).reset_index(drop = True)

    ref_petprod_2.loc[ref_petprod_2['fuel_code'] == 'Crude oil & NGL', 'fuel_code'] = 'Petroleum products'
    ref_petprod_2.loc[ref_petprod_2['item_code_new'] == 'Refining', 'item_code_new'] = 'Domestic refining'
    ref_petprod_2.loc[ref_petprod_2['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_petprod_2.loc[ref_petprod_2['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    ref_petprod_2.loc[ref_petprod_2['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    ref_petprod_2.loc[ref_petprod_2['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    ref_petprod_2 = ref_petprod_2[ref_petprod_2['item_code_new'].isin(fuel_final_ref)].reset_index(drop = True)

    ref_petprod_2['item_code_new'] = pd.Categorical(
        ref_petprod_2['item_code_new'], 
        categories = fuel_final_ref, 
        ordered = True)

    ref_petprod_2 = ref_petprod_2.sort_values('item_code_new').reset_index(drop = True)

    ref_petprod_2_rows = ref_petprod_2.shape[0]
    ref_petprod_2_cols = ref_petprod_2.shape[1]

    #######################

    netz_petprod_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(fuel_vector_ref)) &
                                        (EGEDA_years_netzero['fuel_code'] == '7_petroleum_products')].copy()\
                                            [['fuel_code', 'item_code_new'] + col_chart_years]\
                                                .reset_index(drop = True)
    
    netz_petprod_1['fuel_code'].replace({'7_petroleum_products': 'Petroleum products'}, inplace=True)

    petprod_bunkers = netz_petprod_1[netz_petprod_1['item_code_new'].isin(['4_international_marine_bunkers',
                                                                         '5_international_aviation_bunkers'])]\
                                                                             .groupby(['fuel_code']).sum().assign(fuel_code = 'Petroleum products',
                                                                                                                  item_code_new = 'Bunkers')

    netz_petprod_2 = netz_petprod_1.append([petprod_bunkers]).reset_index(drop = True)

    if 'Refining' in list(netz_crudecons_1['item_code_new']):
        refining_grab = netz_crudecons_1[netz_crudecons_1['item_code_new'] == 'Refining'][['fuel_code', 'item_code_new'] + col_chart_years].copy()

    else:
        refining_series = ['Crude oil & NGL', 'Refining'] + [0] * 7
        refining_grab = pd.Series(refining_series, index = netz_petprod_1.columns)

    netz_petprod_2 = netz_petprod_2.append(refining_grab, ignore_index = True).reset_index(drop = True)

    netz_petprod_2.loc[netz_petprod_2['fuel_code'] == 'Crude oil & NGL', 'fuel_code'] = 'Petroleum products'
    netz_petprod_2.loc[netz_petprod_2['item_code_new'] == 'Refining', 'item_code_new'] = 'Domestic refining'
    netz_petprod_2.loc[netz_petprod_2['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_petprod_2.loc[netz_petprod_2['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'
    netz_petprod_2.loc[netz_petprod_2['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock change'
    netz_petprod_2.loc[netz_petprod_2['item_code_new'] == '7_total_primary_energy_supply', 'item_code_new'] = 'Total primary energy supply'

    netz_petprod_2 = netz_petprod_2[netz_petprod_2['item_code_new'].isin(fuel_final_ref)].reset_index(drop = True)

    netz_petprod_2['item_code_new'] = pd.Categorical(
        netz_petprod_2['item_code_new'], 
        categories = fuel_final_ref, 
        ordered = True)

    netz_petprod_2 = netz_petprod_2.sort_values('item_code_new').reset_index(drop = True)

    netz_petprod_2_rows = netz_petprod_2.shape[0]
    netz_petprod_2_cols = netz_petprod_2.shape[1]

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

    # REFERENCE
    ref_elec_1 = ref_bld_2[ref_bld_2['fuel_code'] == 'Electricity'].copy()\
        .append(ref_ind_2[ref_ind_2['fuel_code'] == 'Electricity'].copy())\
            .append(ref_trn_1[ref_trn_1['fuel_code'] == 'Electricity'].copy())\
                .append(ref_ag_1[ref_ag_1['fuel_code'] == 'Electricity'].copy()).reset_index(drop = True)

    ref_elec_nons1 = EGEDA_years_reference[(EGEDA_years_reference['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                           (EGEDA_years_reference['fuel_code'] == '17_electricity') & 
                                           (EGEDA_years_reference['economy'] == economy)].copy()\
                                               [['fuel_code', 'item_code_new'] + list(EGEDA_years_reference.loc[:,'2000':'2050'])]

    ref_elec_1 = ref_elec_1.append(ref_elec_nons1.copy()).reset_index(drop = True)

    ref_elec_own1 = ref_ownuse_1[ref_ownuse_1['FUEL'] == 'Electricity'].copy()\
        .rename(columns = {'FUEL': 'fuel_code', 'Sector': 'item_code_new'}).reset_index(drop = True)

    ref_elec_1 = ref_elec_1.append(ref_elec_own1.copy()).replace(np.nan, 0).reset_index(drop = True)

    ref_elec_1.loc[ref_elec_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_elec_1.loc[ref_elec_1['item_code_new'] == '16_x_buildings', 'item_code_new'] = 'Buildings'
    ref_elec_1.loc[ref_elec_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_elec_1.loc[ref_elec_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    ref_elec_1.loc[ref_elec_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    ref_elec_1.loc['Total'] = ref_elec_1.sum(numeric_only = True)

    ref_elec_1.loc['Total', 'fuel_code'] = 'Electricity'
    ref_elec_1.loc['Total', 'item_code_new'] = 'Total'

    ref_elec_1 = ref_elec_1.copy().reset_index(drop = True)

    ref_elec_1_rows = ref_elec_1.shape[0]
    ref_elec_1_cols = ref_elec_1.shape[1]

    # CARBON NEUTRALITY
    netz_elec_1 = netz_bld_2[netz_bld_2['fuel_code'] == 'Electricity'].copy()\
        .append(netz_ind_2[netz_ind_2['fuel_code'] == 'Electricity'].copy())\
            .append(netz_trn_1[netz_trn_1['fuel_code'] == 'Electricity'].copy())\
                .append(netz_ag_1[netz_ag_1['fuel_code'] == 'Electricity'].copy()).reset_index(drop = True)

    netz_elec_nons1 = EGEDA_years_netzero[(EGEDA_years_netzero['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                           (EGEDA_years_netzero['fuel_code'] == '17_electricity') & 
                                           (EGEDA_years_netzero['economy'] == economy)].copy()\
                                               [['fuel_code', 'item_code_new'] + list(EGEDA_years_netzero.loc[:,'2000':'2050'])]

    netz_elec_1 = netz_elec_1.append(netz_elec_nons1.copy()).reset_index(drop = True)

    netz_elec_own1 = netz_ownuse_1[netz_ownuse_1['FUEL'] == 'Electricity'].copy()\
        .rename(columns = {'FUEL': 'fuel_code', 'Sector': 'item_code_new'}).reset_index(drop = True)

    netz_elec_1 = netz_elec_1.append(netz_elec_own1.copy()).replace(np.nan, 0).reset_index(drop = True)

    netz_elec_1.loc[netz_elec_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_elec_1.loc[netz_elec_1['item_code_new'] == '16_x_buildings', 'item_code_new'] = 'Buildings'
    netz_elec_1.loc[netz_elec_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_elec_1.loc[netz_elec_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    netz_elec_1.loc[netz_elec_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    netz_elec_1.loc['Total'] = netz_elec_1.sum(numeric_only = True)

    netz_elec_1.loc['Total', 'fuel_code'] = 'Electricity'
    netz_elec_1.loc['Total', 'item_code_new'] = 'Total'

    netz_elec_1 = netz_elec_1.copy().reset_index(drop = True)

    netz_elec_1_rows = netz_elec_1.shape[0]
    netz_elec_1_cols = netz_elec_1.shape[1]

    # Building some waterfall data frames

    if economy in ['01_AUS', '02_BD', '03_CDA', '04_CHL', '05_PRC', '06_HKC',
                   '07_INA', '08_JPN', '09_ROK', '10_MAS', '11_MEX', '12_NZ',
                   '13_PNG', '14_PE', '15_RP', '16_RUS', '17_SIN', '18_CT', '19_THA',
                   '20_USA', '21_VN', 'APEC']:

        # Some key variable for the dataframe constructions to populate dataframe
        ref_emissions_2018 = emiss_total_1.loc[0, '2018']
        ref_emissions_2050 = emiss_total_1.loc[0, '2050']
        netz_emissions_2018 = emiss_total_1.loc[1, '2018']
        netz_emissions_2050 = emiss_total_1.loc[1, '2050']
        pop_growth = (macro_1.loc[macro_1['Series'] == 'Population', '2050'] / macro_1.loc[macro_1['Series'] == 'Population', '2018']).to_numpy()
        gdp_pc_growth = (macro_1.loc[macro_1['Series'] == 'GDP per capita', '2050'] / macro_1.loc[macro_1['Series'] == 'GDP per capita', '2018']).to_numpy()
        ref_ei_growth = (ref_enint_sup3.loc[ref_enint_sup3['Series'] == 'Reference', '2050'] / ref_enint_sup3.loc[ref_enint_sup3['Series'] == 'Reference', '2018']).to_numpy()
        ref_co2i_growth = (ref_co2int_2.loc[ref_co2int_2['item_code_new'] == 'CO2 intensity', '2050'] / ref_co2int_2.loc[ref_co2int_2['item_code_new'] == 'CO2 intensity', '2018']).to_numpy()
        netz_ei_growth = (netz_enint_sup3.loc[netz_enint_sup3['Series'] == 'Carbon Neutrality', '2050'] / netz_enint_sup3.loc[netz_enint_sup3['Series'] == 'Carbon Neutrality', '2018']).to_numpy()
        netz_co2i_growth = (netz_co2int_2.loc[netz_co2int_2['item_code_new'] == 'CO2 intensity', '2050'] / netz_co2int_2.loc[netz_co2int_2['item_code_new'] == 'CO2 intensity', '2018']).to_numpy()

        if (pop_growth >= 1) & (ref_co2i_growth < 1) & (ref_ei_growth < 1):

            ref_kaya_1 = pd.DataFrame(index = [list(range(7))], 
                                    columns = ['Reference', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            ref_kaya_1.loc[0, 'Reference'] = 'initial'
            ref_kaya_1.loc[1, 'Reference'] = 'empty'
            ref_kaya_1.loc[2, 'Reference'] = 'no improve'
            ref_kaya_1.loc[3, 'Reference'] = 'empty'
            ref_kaya_1.loc[4, 'Reference'] = 'no improve'
            ref_kaya_1.loc[5, 'Reference'] = 'improve'
            ref_kaya_1.loc[6, 'Reference'] = 'improve'

            # Emissions 2018 column
            ref_kaya_1.loc[0, 'Emissions 2018'] = ref_emissions_2018
            ref_kaya_1.loc[0, 'Emissions 2050'] = ref_emissions_2050
            
            # Population column (Emissions multiplied by population factor split into two data points)
            ref_kaya_1.loc[1, 'Population'] = ref_emissions_2018
            ref_kaya_1.loc[2, 'Population'] = (ref_emissions_2018 * pop_growth) - ref_emissions_2018

            # GDP per capita column
            ref_kaya_1.loc[1, 'GDP per capita'] = ref_emissions_2018
            ref_kaya_1.loc[3, 'GDP per capita'] = (ref_emissions_2018 * pop_growth) - ref_emissions_2018
            ref_kaya_1.loc[4, 'GDP per capita'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth) - (ref_emissions_2018 * pop_growth)

            # Energy intensity column
            ref_kaya_1.loc[1, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth)
            ref_kaya_1.loc[5, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth)

            # Emissions intensity column
            ref_kaya_1.loc[1, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth * ref_co2i_growth)
            ref_kaya_1.loc[6, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth * ref_co2i_growth)

            ref_kaya_1 = ref_kaya_1.copy().replace(np.nan, 0).reset_index(drop = True)

            ref_kaya_1_rows = ref_kaya_1.shape[0]
            ref_kaya_1_cols = ref_kaya_1.shape[1]

        elif (pop_growth < 1) & (ref_co2i_growth < 1) & (ref_ei_growth < 1):

            ref_kaya_1 = pd.DataFrame(index = [list(range(7))], 
                                    columns = ['Reference', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            ref_kaya_1.loc[0, 'Reference'] = 'initial'
            ref_kaya_1.loc[1, 'Reference'] = 'empty'
            ref_kaya_1.loc[2, 'Reference'] = 'improve'
            ref_kaya_1.loc[3, 'Reference'] = 'empty'
            ref_kaya_1.loc[4, 'Reference'] = 'no improve'
            ref_kaya_1.loc[5, 'Reference'] = 'improve'
            ref_kaya_1.loc[6, 'Reference'] = 'improve'

            # Emissions 2018 column
            ref_kaya_1.loc[0, 'Emissions 2018'] = ref_emissions_2018
            ref_kaya_1.loc[0, 'Emissions 2050'] = ref_emissions_2050
            
            # Population column (Emissions multiplied by population factor split into two data points)
            ref_kaya_1.loc[1, 'Population'] = ref_emissions_2018 * pop_growth
            ref_kaya_1.loc[2, 'Population'] = ref_emissions_2018 - (ref_emissions_2018 * pop_growth)  

            # GDP per capita column
            ref_kaya_1.loc[1, 'GDP per capita'] = ref_emissions_2018 * pop_growth
            # ref_kaya_1.loc[3, 'GDP per capita'] = (ref_emissions_2018 * pop_growth) - ref_emissions_2018
            ref_kaya_1.loc[4, 'GDP per capita'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth) - (ref_emissions_2018 * pop_growth)

            # Energy intensity column
            ref_kaya_1.loc[1, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth)
            ref_kaya_1.loc[5, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth)

            # Emissions intensity column
            ref_kaya_1.loc[1, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth * ref_co2i_growth)
            ref_kaya_1.loc[6, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth * ref_co2i_growth)

            ref_kaya_1 = ref_kaya_1.copy().replace(np.nan, 0).reset_index(drop = True)

            ref_kaya_1_rows = ref_kaya_1.shape[0]
            ref_kaya_1_cols = ref_kaya_1.shape[1]

        elif (pop_growth >= 1) & (ref_co2i_growth >= 1) & (ref_ei_growth < 1):

            ref_kaya_1 = pd.DataFrame(index = [list(range(7))], 
                                    columns = ['Reference', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            ref_kaya_1.loc[0, 'Reference'] = 'initial'
            ref_kaya_1.loc[1, 'Reference'] = 'empty'
            ref_kaya_1.loc[2, 'Reference'] = 'no improve'
            ref_kaya_1.loc[3, 'Reference'] = 'empty'
            ref_kaya_1.loc[4, 'Reference'] = 'no improve'
            ref_kaya_1.loc[5, 'Reference'] = 'improve'
            ref_kaya_1.loc[6, 'Reference'] = 'no improve'

            # Emissions 2018 column
            ref_kaya_1.loc[0, 'Emissions 2018'] = ref_emissions_2018
            ref_kaya_1.loc[0, 'Emissions 2050'] = ref_emissions_2050
            
            # Population column (Emissions multiplied by population factor split into two data points)
            ref_kaya_1.loc[1, 'Population'] = ref_emissions_2018
            ref_kaya_1.loc[2, 'Population'] = (ref_emissions_2018 * pop_growth) - ref_emissions_2018

            # GDP per capita column
            ref_kaya_1.loc[1, 'GDP per capita'] = ref_emissions_2018
            ref_kaya_1.loc[3, 'GDP per capita'] = (ref_emissions_2018 * pop_growth) - ref_emissions_2018
            ref_kaya_1.loc[4, 'GDP per capita'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth) - (ref_emissions_2018 * pop_growth)

            # Energy intensity column
            ref_kaya_1.loc[1, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth)
            ref_kaya_1.loc[5, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth)

            # Emissions intensity column
            ref_kaya_1.loc[1, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth) 
            ref_kaya_1.loc[6, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth * ref_co2i_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth) 

            ref_kaya_1 = ref_kaya_1.copy().replace(np.nan, 0).reset_index(drop = True)

            ref_kaya_1_rows = ref_kaya_1.shape[0]
            ref_kaya_1_cols = ref_kaya_1.shape[1]

        elif (pop_growth < 1) & (ref_co2i_growth >= 1) & (ref_ei_growth < 1):

            ref_kaya_1 = pd.DataFrame(index = [list(range(7))], 
                                    columns = ['Reference', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            ref_kaya_1.loc[0, 'Reference'] = 'initial'
            ref_kaya_1.loc[1, 'Reference'] = 'empty'
            ref_kaya_1.loc[2, 'Reference'] = 'improve'
            ref_kaya_1.loc[3, 'Reference'] = 'empty'
            ref_kaya_1.loc[4, 'Reference'] = 'no improve'
            ref_kaya_1.loc[5, 'Reference'] = 'improve'
            ref_kaya_1.loc[6, 'Reference'] = 'no improve'

            # Emissions 2018 column
            ref_kaya_1.loc[0, 'Emissions 2018'] = ref_emissions_2018
            ref_kaya_1.loc[0, 'Emissions 2050'] = ref_emissions_2050
            
            # Population column (Emissions multiplied by population factor split into two data points)
            ref_kaya_1.loc[1, 'Population'] = ref_emissions_2018 * pop_growth
            ref_kaya_1.loc[2, 'Population'] = ref_emissions_2018 - (ref_emissions_2018 * pop_growth)  

            # GDP per capita column
            ref_kaya_1.loc[1, 'GDP per capita'] = ref_emissions_2018 * pop_growth
            # ref_kaya_1.loc[3, 'GDP per capita'] = (ref_emissions_2018 * pop_growth) - ref_emissions_2018
            ref_kaya_1.loc[4, 'GDP per capita'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth) - (ref_emissions_2018 * pop_growth)

            # Energy intensity column
            ref_kaya_1.loc[1, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth)
            ref_kaya_1.loc[5, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth)

            # Emissions intensity column
            ref_kaya_1.loc[1, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth) 
            ref_kaya_1.loc[6, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth * ref_co2i_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth) 

            ref_kaya_1 = ref_kaya_1.copy().replace(np.nan, 0).reset_index(drop = True)

            ref_kaya_1_rows = ref_kaya_1.shape[0]
            ref_kaya_1_cols = ref_kaya_1.shape[1]

        elif (pop_growth >= 1) & (ref_co2i_growth >= 1) & (ref_ei_growth >= 1):

            ref_kaya_1 = pd.DataFrame(index = [list(range(7))], 
                                    columns = ['Reference', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            ref_kaya_1.loc[0, 'Reference'] = 'initial'
            ref_kaya_1.loc[1, 'Reference'] = 'empty'
            ref_kaya_1.loc[2, 'Reference'] = 'no improve'
            ref_kaya_1.loc[3, 'Reference'] = 'empty'
            ref_kaya_1.loc[4, 'Reference'] = 'no improve'
            ref_kaya_1.loc[5, 'Reference'] = 'no improve'
            ref_kaya_1.loc[6, 'Reference'] = 'no improve'

            # Emissions 2018 column
            ref_kaya_1.loc[0, 'Emissions 2018'] = ref_emissions_2018
            ref_kaya_1.loc[0, 'Emissions 2050'] = ref_emissions_2050
            
            # Population column (Emissions multiplied by population factor split into two data points)
            ref_kaya_1.loc[1, 'Population'] = ref_emissions_2018
            ref_kaya_1.loc[2, 'Population'] = (ref_emissions_2018 * pop_growth) - ref_emissions_2018   

            # GDP per capita column
            ref_kaya_1.loc[1, 'GDP per capita'] = ref_emissions_2018
            ref_kaya_1.loc[3, 'GDP per capita'] = (ref_emissions_2018 * pop_growth) - ref_emissions_2018
            ref_kaya_1.loc[4, 'GDP per capita'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth) - (ref_emissions_2018 * pop_growth)

            # Energy intensity column
            ref_kaya_1.loc[1, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth)
            ref_kaya_1.loc[5, 'Energy intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth)

            # Emissions intensity column
            ref_kaya_1.loc[1, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth) 
            ref_kaya_1.loc[6, 'Emissions intensity'] = (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth * ref_co2i_growth) - (ref_emissions_2018 * pop_growth * gdp_pc_growth * ref_ei_growth) 

            ref_kaya_1 = ref_kaya_1.copy().replace(np.nan, 0).reset_index(drop = True)

            ref_kaya_1_rows = ref_kaya_1.shape[0]
            ref_kaya_1_cols = ref_kaya_1.shape[1]

        else:
            pass

        # Now if statements for Carbon neutrality data frame builds

        if (pop_growth >= 1) & (netz_co2i_growth < 1):

            netz_kaya_1 = pd.DataFrame(index = [list(range(7))], 
                                    columns = ['Carbon Neutrality', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            netz_kaya_1.loc[0, 'Carbon Neutrality'] = 'initial'
            netz_kaya_1.loc[1, 'Carbon Neutrality'] = 'empty'
            netz_kaya_1.loc[2, 'Carbon Neutrality'] = 'no improve'
            netz_kaya_1.loc[3, 'Carbon Neutrality'] = 'empty'
            netz_kaya_1.loc[4, 'Carbon Neutrality'] = 'no improve'
            netz_kaya_1.loc[5, 'Carbon Neutrality'] = 'improve'
            netz_kaya_1.loc[6, 'Carbon Neutrality'] = 'improve'

            # Emissions 2018 column
            netz_kaya_1.loc[0, 'Emissions 2018'] = netz_emissions_2018
            netz_kaya_1.loc[0, 'Emissions 2050'] = netz_emissions_2050
            
            # Population column (Emissions multiplied by population factor split into two data points)
            netz_kaya_1.loc[1, 'Population'] = netz_emissions_2018
            netz_kaya_1.loc[2, 'Population'] = (netz_emissions_2018 * pop_growth) - netz_emissions_2018

            # GDP per capita column
            netz_kaya_1.loc[1, 'GDP per capita'] = netz_emissions_2018
            netz_kaya_1.loc[3, 'GDP per capita'] = (netz_emissions_2018 * pop_growth) - netz_emissions_2018
            netz_kaya_1.loc[4, 'GDP per capita'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth) - (netz_emissions_2018 * pop_growth)

            # Energy intensity column
            netz_kaya_1.loc[1, 'Energy intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth)
            netz_kaya_1.loc[5, 'Energy intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth) - (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth)

            # Emissions intensity column
            netz_kaya_1.loc[1, 'Emissions intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth * netz_co2i_growth)
            netz_kaya_1.loc[6, 'Emissions intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth) - (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth * netz_co2i_growth)

            netz_kaya_1 = netz_kaya_1.copy().replace(np.nan, 0).reset_index(drop = True)

            netz_kaya_1_rows = netz_kaya_1.shape[0]
            netz_kaya_1_cols = netz_kaya_1.shape[1]

        elif (pop_growth < 1) & (netz_co2i_growth < 1):

            netz_kaya_1 = pd.DataFrame(index = [list(range(7))], 
                                    columns = ['Carbon Neutrality', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            netz_kaya_1.loc[0, 'Carbon Neutrality'] = 'initial'
            netz_kaya_1.loc[1, 'Carbon Neutrality'] = 'empty'
            netz_kaya_1.loc[2, 'Carbon Neutrality'] = 'improve'
            netz_kaya_1.loc[3, 'Carbon Neutrality'] = 'empty'
            netz_kaya_1.loc[4, 'Carbon Neutrality'] = 'no improve'
            netz_kaya_1.loc[5, 'Carbon Neutrality'] = 'improve'
            netz_kaya_1.loc[6, 'Carbon Neutrality'] = 'improve'

            # Emissions 2018 column
            netz_kaya_1.loc[0, 'Emissions 2018'] = netz_emissions_2018
            netz_kaya_1.loc[0, 'Emissions 2050'] = netz_emissions_2050
            
            # Population column (Emissions multiplied by population factor split into two data points)
            netz_kaya_1.loc[1, 'Population'] = netz_emissions_2018 * pop_growth
            netz_kaya_1.loc[2, 'Population'] = netz_emissions_2018 - (netz_emissions_2018 * pop_growth)  

            # GDP per capita column
            netz_kaya_1.loc[1, 'GDP per capita'] = netz_emissions_2018 * pop_growth
            # netz_kaya_1.loc[3, 'GDP per capita'] = (netz_emissions_2018 * pop_growth) - netz_emissions_2018
            netz_kaya_1.loc[4, 'GDP per capita'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth) - (netz_emissions_2018 * pop_growth)

            # Energy intensity column
            netz_kaya_1.loc[1, 'Energy intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth)
            netz_kaya_1.loc[5, 'Energy intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth) - (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth)

            # Emissions intensity column
            netz_kaya_1.loc[1, 'Emissions intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth * netz_co2i_growth)
            netz_kaya_1.loc[6, 'Emissions intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth) - (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth * netz_co2i_growth)

            netz_kaya_1 = netz_kaya_1.copy().replace(np.nan, 0).reset_index(drop = True)

            netz_kaya_1_rows = netz_kaya_1.shape[0]
            netz_kaya_1_cols = netz_kaya_1.shape[1]

        elif (pop_growth >= 1) & (netz_co2i_growth >= 1):

            netz_kaya_1 = pd.DataFrame(index = [list(range(7))], 
                                    columns = ['Carbon Neutrality', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            netz_kaya_1.loc[0, 'Carbon Neutrality'] = 'initial'
            netz_kaya_1.loc[1, 'Carbon Neutrality'] = 'empty'
            netz_kaya_1.loc[2, 'Carbon Neutrality'] = 'no improve'
            netz_kaya_1.loc[3, 'Carbon Neutrality'] = 'empty'
            netz_kaya_1.loc[4, 'Carbon Neutrality'] = 'no improve'
            netz_kaya_1.loc[5, 'Carbon Neutrality'] = 'improve'
            netz_kaya_1.loc[6, 'Carbon Neutrality'] = 'no improve'

            # Emissions 2018 column
            netz_kaya_1.loc[0, 'Emissions 2018'] = netz_emissions_2018
            netz_kaya_1.loc[0, 'Emissions 2050'] = netz_emissions_2050
            
            # Population column (Emissions multiplied by population factor split into two data points)
            netz_kaya_1.loc[1, 'Population'] = netz_emissions_2018
            netz_kaya_1.loc[2, 'Population'] = (netz_emissions_2018 * pop_growth) - netz_emissions_2018

            # GDP per capita column
            netz_kaya_1.loc[1, 'GDP per capita'] = netz_emissions_2018
            netz_kaya_1.loc[3, 'GDP per capita'] = (netz_emissions_2018 * pop_growth) - netz_emissions_2018
            netz_kaya_1.loc[4, 'GDP per capita'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth) - (netz_emissions_2018 * pop_growth)

            # Energy intensity column
            netz_kaya_1.loc[1, 'Energy intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth)
            netz_kaya_1.loc[5, 'Energy intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth) - (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth)

            # Emissions intensity column
            netz_kaya_1.loc[1, 'Emissions intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth) 
            netz_kaya_1.loc[6, 'Emissions intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth * netz_co2i_growth) - (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth) 

            netz_kaya_1 = netz_kaya_1.copy().replace(np.nan, 0).reset_index(drop = True)

            netz_kaya_1_rows = netz_kaya_1.shape[0]
            netz_kaya_1_cols = netz_kaya_1.shape[1]

        elif (pop_growth < 1) & (netz_co2i_growth >= 1):

            netz_kaya_1 = pd.DataFrame(index = [list(range(7))], 
                                    columns = ['Carbon Neutrality', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            netz_kaya_1.loc[0, 'Carbon Neutrality'] = 'initial'
            netz_kaya_1.loc[1, 'Carbon Neutrality'] = 'empty'
            netz_kaya_1.loc[2, 'Carbon Neutrality'] = 'improve'
            netz_kaya_1.loc[3, 'Carbon Neutrality'] = 'empty'
            netz_kaya_1.loc[4, 'Carbon Neutrality'] = 'no improve'
            netz_kaya_1.loc[5, 'Carbon Neutrality'] = 'improve'
            netz_kaya_1.loc[6, 'Carbon Neutrality'] = 'no improve'

            # Emissions 2018 column
            netz_kaya_1.loc[0, 'Emissions 2018'] = netz_emissions_2018
            netz_kaya_1.loc[0, 'Emissions 2050'] = netz_emissions_2050
            
            # Population column (Emissions multiplied by population factor split into two data points)
            netz_kaya_1.loc[1, 'Population'] = netz_emissions_2018 * pop_growth
            netz_kaya_1.loc[2, 'Population'] = netz_emissions_2018 - (netz_emissions_2018 * pop_growth)  

            # GDP per capita column
            netz_kaya_1.loc[1, 'GDP per capita'] = netz_emissions_2018 * pop_growth
            # netz_kaya_1.loc[3, 'GDP per capita'] = (netz_emissions_2018 * pop_growth) - netz_emissions_2018
            netz_kaya_1.loc[4, 'GDP per capita'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth) - (netz_emissions_2018 * pop_growth)

            # Energy intensity column
            netz_kaya_1.loc[1, 'Energy intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth)
            netz_kaya_1.loc[5, 'Energy intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth) - (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth)

            # Emissions intensity column
            netz_kaya_1.loc[1, 'Emissions intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth) 
            netz_kaya_1.loc[6, 'Emissions intensity'] = (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth * netz_co2i_growth) - (netz_emissions_2018 * pop_growth * gdp_pc_growth * netz_ei_growth) 

            netz_kaya_1 = netz_kaya_1.copy().replace(np.nan, 0).reset_index(drop = True)

            netz_kaya_1_rows = netz_kaya_1.shape[0]
            netz_kaya_1_cols = netz_kaya_1.shape[1]

        else:
            pass

    else:
        ref_kaya_1 = pd.DataFrame()
        ref_kaya_1_rows = ref_kaya_1.shape[0]
        ref_kaya_1_cols = ref_kaya_1.shape[1]

        netz_kaya_1 = pd.DataFrame()
        netz_kaya_1_rows = netz_kaya_1.shape[0]
        netz_kaya_1_cols = netz_kaya_1.shape[1]

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

    ref_tpes_pGDP = ['TPES per GDP (MJ per 2018 USD PPP)'] + list(ref_tpes_calcs.iloc[0, 1:] / ref_tpes_calcs.iloc[1, 1:])
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

    ref_tfec_pGDP = ['Final energy intensity (MJ per 2018 USD PPP)'] + list(ref_tfec_calcs.iloc[1, 1:] / ref_tfec_calcs.iloc[2, 1:])
    ref_tfec_pGDP_series = pd.Series(ref_tfec_pGDP, index = ref_tfec_1.columns)

    ref_tfec_1 = ref_tfec_1.append([ref_tfec_pc_series, ref_tfec_pGDP_series], ignore_index = True).reset_index(drop = True)

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
    ## OUTPUT FUEL
    # Electricity
    # Heat

    ### REFINERIES
    ## INPUT FUEL
    # Crude oil
    ## OUTPUT FUEL
    # Refined products

    ### ENERGY INDUSTRY OWN-USE
    ## INPUT FUEL
    # Coal 
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat 
    # Other
    
    ### DISTRIBUTION LOSSES
    # Coal
    # Oil
    # Gas
    # Electricity
    # Heat

    ### TRANSFERS

    ### STATISTICAL DISCREPANCIES (only historical in 7th)

    ######################################################

    ### DEMAND
    ## FED By Sector
    # Agriculture and non-specified (split)
    # Buildings
    # Industry
    # Transport (domestic)
    # Non-energy

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

    ## INDUSTRY
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    ## TRANSPORT
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    ## NON-ENERGY
    # Coal
    # Oil
    # Gas

    ###############################################################

    #### ELECTRICITY
    ### CAPACITY
    ## COAL
    # Subcritical
    # Supercritical and USC
    # Advanced USC and integrated gasification combined cycle
    # Coal CHP
    # Coal with CCS
    # 
    ## OIL
    # Non-CHP
    # Oil CHP
    #
    ## GAS
    # Combined cycle
    # Turbine
    # Gas CHP
    # Gas with CCS
    # 
    ## Nuclear
    # 
    ## RENEWABLES
    # Hydro
    # Wind 
    # Solar
    # Bioenergy
    # Geothermal
    # Other renewables
    # 
    ## OTHER
    ## STORAGE
    # 
    ###################################
    ### GENERATION 
    ## COAL
    # Subcritical
    # Supercritical and USC
    # Advanced USC and integrated gasification combined cycle
    # Coal CHP
    # Coal with CCS
    # 
    ## OIL
    # Non-CHP
    # Oil CHP
    #
    ## GAS
    # Combined cycle
    # Turbine
    # Gas CHP
    # Gas with CCS
    # 
    ## Nuclear
    # 
    ## RENEWABLES
    # Hydro
    # Wind 
    # Solar
    # Bioenergy
    # Geothermal
    # Other renewables
    # 
    ## OTHER
    ## STORAGE

    ######################################
    ### EMISSIONS
    ## BY FUEL
    # Coal
    # Oil
    # Gas

    ## BY SECTOR
    # Power
    # Own-use and losses
    # Agriculture 
    # Buildings
    # Industry
    # Transport (domestic)
    # Non-energy
    # Non-specified


    # Define directory to save charts and tables workbook
    script_dir = './results/'
    results_dir = os.path.join(script_dir, economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)
        
    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_charts_' + day_month_year + '.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    # Insert the various dataframes into different sheets of the workbook
    # REFERENCE and NETZERO

    # Macro
    macro_1.to_excel(writer, sheet_name = 'Macro', index = False, startrow = chart_height)

        
    writer.save()

print('Bling blang blaow, you have some charts now')

