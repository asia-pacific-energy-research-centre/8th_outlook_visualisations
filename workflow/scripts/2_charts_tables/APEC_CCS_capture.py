# All portions of FED, Supply and Transformation charts in one script

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

heat_prod_tech = ['Coal', 'Lignite', 'Oil', 'Gas', 'Gas CCS', 'Nuclear', 'Biomass', 'Waste', 'Non-specified', 'Heat only units', 'Total']

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


pow_capacity_agg = ['Coal', 'Coal CCS', 'Gas', 'Gas CCS', 'Oil', 'Nuclear', 'Hydro', 'Bio', 'Wind', 'Solar', 'Geothermal', 'Waste', 'Storage', 'Other']
pow_capacity_agg2 = ['Coal', 'Coal CCS', 'Lignite', 'Gas', 'Gas CCS', 'Oil', 'Nuclear', 'Hydro', 'Bio', 'Wind', 
                     'Solar', 'Geothermal', 'Waste', 'Storage', 'Other']

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
                                                                '15_solid_biomass': 'Bio', 
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

# CCS dataframe builds

# Reference dataframe for power and hydrogen energy use

# REF
ref_ccs_power = ref_power_df1[(ref_power_df1['TECHNOLOGY'].isin(['POW_CCGT_CCS_PP', 'POW_COAL_CCS_PP'])) &
                                (ref_power_df1['Sheet_energy'] == 'UseByTechnology')].copy().reset_index(drop = True)

ref_ccs_hyd = ref_refownsup_df1[(ref_refownsup_df1['TECHNOLOGY'].isin(['HYD_ng_smr_ccs', 'HYD_coal_gas_ccs', 'HYD_ng_smr_ccs_export'])) &
                                (ref_refownsup_df1['Sheet_energy'] == 'UseByTechnology') &
                                (ref_refownsup_df1['FUEL'].isin(['8_1_natural_gas', '1_1_coking_coal']))].copy().reset_index(drop = True)

ref_ccs_1 = ref_ccs_power.append(ref_ccs_hyd).copy().reset_index(drop = True)

ref_ccs_1_rows = ref_ccs_1.shape[0]
ref_ccs_1_cols = ref_ccs_1.shape[1]

# CN

cn_ccs_power = netz_power_df1[(netz_power_df1['TECHNOLOGY'].isin(['POW_CCGT_CCS_PP', 'POW_COAL_CCS_PP'])) &
                              (netz_power_df1['Sheet_energy'] == 'UseByTechnology')].copy().reset_index(drop = True)

cn_ccs_hyd = netz_refownsup_df1[(netz_refownsup_df1['TECHNOLOGY'].isin(['HYD_ng_smr_ccs', 'HYD_coal_gas_ccs', 'HYD_ng_smr_ccs_export'])) &
                                (netz_refownsup_df1['Sheet_energy'] == 'UseByTechnology') &
                                (netz_refownsup_df1['FUEL'].isin(['8_1_natural_gas', '1_1_coking_coal']))].copy().reset_index(drop = True)

cn_ccs_1 = cn_ccs_power.append(cn_ccs_hyd).copy().reset_index(drop = True)

cn_ccs_1_rows = cn_ccs_1.shape[0]
cn_ccs_1_cols = cn_ccs_1.shape[1]

# Df builds are complete

##############################################################################################################################

# Define directory to save charts and tables workbook
script_dir = './results/APEC_ccs'
results_dir = os.path.join(script_dir)
if not os.path.isdir(results_dir):
    os.makedirs(results_dir)
    
# Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
writer = pd.ExcelWriter(results_dir + '/APEC_ccs.xlsx', engine = 'xlsxwriter')
workbook = writer.book
pandas.io.formats.excel.ExcelFormatter.header_style = None

# Insert the various dataframes into different sheets of the workbook
# REFERENCE and NETZERO

ref_ccs_1.to_excel(writer, sheet_name = 'CCS_power_hyd', index = False, startrow = chart_height)
cn_ccs_1.to_excel(writer, sheet_name = 'CCS_power_hyd', index = False, startrow = chart_height + ref_ccs_1_rows + 3)

# Access the workbook and first sheet with data from df1
ref_worksheet1 = writer.sheets['CCS_power_hyd']

# Comma format and header format        
space_format = workbook.add_format({'num_format': '# ### ### ##0.0;-# ### ### ##0.0;-'})
percentage_format = workbook.add_format({'num_format': '0.0%'})
header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
cell_format1 = workbook.add_format({'bold': True})
cell_format2 = workbook.add_format({'font_size': 9})
    
# Apply comma format and header format to relevant data rows
ref_worksheet1.set_column(1, ref_ccs_1_cols + 1, None, space_format)
ref_worksheet1.set_row(chart_height, None, header_format)
ref_worksheet1.set_row(chart_height + ref_ccs_1_rows + 3, None, header_format)
ref_worksheet1.write(0, 0, economy + ' CCS facilities in power and hydrogen sector use of fossil fuels, REF and CN', cell_format1)
ref_worksheet1.write(1, 0, 'Units: Petajoules', cell_format2)   

writer.save()

print('Bling blang blaow, you have some charts now')

