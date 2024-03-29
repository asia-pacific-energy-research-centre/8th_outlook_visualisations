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
    non_zero = (ref_imports_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_imports_1 = ref_imports_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (ref_exports_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_exports_1 = ref_exports_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (ref_bunkers_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_bunkers_1 = ref_bunkers_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (ref_bunkers_2.loc[:,'2000':] != 0).any(axis = 1)
    ref_bunkers_2 = ref_bunkers_2.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_imports_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_imports_1 = netz_imports_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_exports_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_exports_1 = netz_exports_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_bunkers_1.loc[:,'2000':] != 0).any(axis = 1)
    netz_bunkers_1 = netz_bunkers_1.loc[non_zero].reset_index(drop = True)

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
    non_zero = (netz_bunkers_2.loc[:,'2000':] != 0).any(axis = 1)
    netz_bunkers_2 = netz_bunkers_2.loc[non_zero].reset_index(drop = True)

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
        [['FUEL', 'TECHNOLOGY'] + list(ref_pow_use_1.loc[:, '2017':])].reset_index(drop = True)

    ref_pow_use_2 = ref_pow_use_2[ref_pow_use_2['FUEL'].isin(use_agg_fuels_1)].copy().set_index('FUEL').reset_index()

    ref_pow_use_2 = ref_pow_use_2.groupby('FUEL').sum().reset_index()
    ref_pow_use_2['Transformation'] = 'Input fuel'

    #################################################################################
    historical_input = EGEDA_hist_power[EGEDA_hist_power['economy'] == economy].copy().\
        iloc[:,:-2][['FUEL', 'Transformation'] + list(EGEDA_hist_power.loc[:, '2000':'2016'])]

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

    # Generation of electricity by tech dataframe (with the above aggregations added)

    ref_elecgen_2 = ref_elecgen_1.append([coal_pp2, coal_ccs_pp, lignite_pp2, oil_pp, gas_pp, gas_ccs_pp, storage_pp, nuclear_pp,\
        bio_pp, geo_pp, waste_pp, other_pp, hydro_pp, misc, solar_pp, wind_pp])\
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

    # Get rid of zero rows
    non_zero = (ref_powcap_1.loc[:,'2018':] != 0).any(axis = 1)
    ref_powcap_1 = ref_powcap_1.loc[non_zero].reset_index(drop = True)

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

    ref_owncoal_1 = ref_owncoal_1.copy().groupby(['FUEL', 'Sector']).sum().reset_index()

    hist_owncoal = EGEDA_hist_owncoal[EGEDA_hist_owncoal['economy'] == economy].copy()\
        .iloc[:,:-2][['FUEL', 'Sector'] + list(EGEDA_hist_owncoal.loc[:, '2000':'2016'])]

    ref_owncoal_1 = hist_owncoal.merge(ref_owncoal_1, how = 'right', on = ['FUEL', 'Sector']).replace(np.nan, 0)

    ref_owncoal_1 = ref_owncoal_1[['FUEL', 'Sector'] + list(ref_owncoal_1.loc[:, '2000':'2050'])]

    # Get rid of zero rows
    non_zero = (ref_owncoal_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_owncoal_1 = ref_owncoal_1.loc[non_zero].reset_index(drop = True)

    ref_owncoal_1_rows = ref_owncoal_1.shape[0]
    ref_owncoal_1_cols = ref_owncoal_1.shape[1]

    #################################################################

    ref_ownuse_1 = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(own_use_fuels)].reset_index(drop = True)

    #################################################################################
    historical_input = EGEDA_hist_own[EGEDA_hist_own['economy'] == economy].copy().\
        iloc[:,:-2][['FUEL', 'Sector'] + list(EGEDA_hist_own.loc[:, '2000':'2016'])]

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
    comb_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(combination_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Heat only units')
    nons_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(nons_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Non-specified')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    ref_heatgen_2 = ref_heatgen_1.append([coal_hp, lignite_hp, oil_hp, gas_hp, gas_ccs_hp, nuclear_hp, bio_hp, waste_hp, comb_hp, nons_hp])\
        [['TECHNOLOGY'] + list(ref_heatgen_1.loc[:, '2017':'2050'])].reset_index(drop = True)                                                                                                    

    ref_heatgen_2['Generation'] = 'Heat'
    ref_heatgen_2 = ref_heatgen_2[['TECHNOLOGY', 'Generation'] + list(ref_heatgen_2.loc[:, '2017':'2050'])] 

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
        iloc[:,:-2][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_heat.loc[:, '2000':'2016'])]

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
        [['FUEL', 'TECHNOLOGY'] + list(netz_pow_use_1.loc[:,'2017':'2050'])].reset_index(drop = True)

    netz_pow_use_2 = netz_pow_use_2[netz_pow_use_2['FUEL'].isin(use_agg_fuels_1)].copy().set_index('FUEL').reset_index()

    netz_pow_use_2 = netz_pow_use_2.groupby('FUEL').sum().reset_index()
    netz_pow_use_2['Transformation'] = 'Input fuel'
    
    #################################################################################
    historical_input = EGEDA_hist_power[EGEDA_hist_power['economy'] == economy].copy().\
        iloc[:,:-2][['FUEL', 'Transformation'] + list(EGEDA_hist_power.loc[:, '2000':'2016'])]

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

    # Generation of electricity by tech dataframe (with the above aggregations added)

    netz_elecgen_2 = netz_elecgen_1.append([coal_pp2, coal_ccs_pp, lignite_pp2, oil_pp, gas_pp, gas_ccs_pp, storage_pp, nuclear_pp,\
        bio_pp, geo_pp, waste_pp, other_pp, hydro_pp, misc, solar_pp, wind_pp])\
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

    # Get rid of zero rows
    non_zero = (netz_powcap_1.loc[:,'2018':] != 0).any(axis = 1)
    netz_powcap_1 = netz_powcap_1.loc[non_zero].reset_index(drop = True)

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
        [['FUEL', 'Sector'] + list(netz_ownuse_1.loc[:,'2017':'2050'])].reset_index(drop = True)

    netz_ownuse_1 = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(own_use_fuels)].reset_index(drop = True)

    #################################################################################
    historical_input = EGEDA_hist_own[EGEDA_hist_own['economy'] == economy].copy().\
        iloc[:,:-2][['FUEL', 'Sector'] + list(EGEDA_hist_own.loc[:, '2000':'2016'])]

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
    comb_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(combination_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Heat only units')
    nons_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(nons_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Non-specified')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    netz_heatgen_2 = netz_heatgen_1.append([coal_hp, lignite_hp, oil_hp, gas_hp, gas_ccs_hp, nuclear_hp, bio_hp, waste_hp, comb_hp, nons_hp])\
        [['TECHNOLOGY'] + list(netz_heatgen_1.loc[:, '2017':'2050'])].reset_index(drop = True)                                                              

    netz_heatgen_2['Generation'] = 'Heat'
    netz_heatgen_2 = netz_heatgen_2[['TECHNOLOGY', 'Generation'] + list(netz_heatgen_2.loc[:, '2017':'2050'])]

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
        iloc[:,:-2][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_heat.loc[:, '2000':'2016'])]

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
    historical_eh = EGEDA_hist_eh[EGEDA_hist_eh['economy'] == economy].copy().iloc[:,1:-2]

    ref_modren_elecheat = historical_eh.merge(ref_modren_elecheat, how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_modren_elecheat = ref_modren_elecheat[['fuel_code', 'item_code_new'] + list(ref_modren_elecheat\
        .loc[:, '2000':'2050'])]

    ref_modren_2 = ref_modren_1.append(ref_modren_elecheat).reset_index(drop = True)

    # NEW EDIT: First slot in 'Total' electricity and heat (including losses and own use)

    # Grab historical for all electricity and heat
    historical_eh2 = EGEDA_hist_eh2[EGEDA_hist_eh2['economy'] == economy].copy().iloc[:, 1:-2]

    ref_all_elecheat = historical_eh2.merge(ref_elecheat, how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
    historical_eh = EGEDA_hist_eh[EGEDA_hist_eh['economy'] == economy].copy().iloc[:,1:-2]

    netz_modren_elecheat = historical_eh.merge(netz_modren_elecheat, how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_modren_elecheat = netz_modren_elecheat[['fuel_code', 'item_code_new'] + list(netz_modren_elecheat\
        .loc[:, '2000':'2050'])]

    netz_modren_2 = netz_modren_1.append(netz_modren_elecheat).reset_index(drop = True)
    # NEW EDIT: First slot in 'Total' electricity and heat (including losses and own use)

    # Grab historical for all electricity and heat
    historical_eh2 = EGEDA_hist_eh2[EGEDA_hist_eh2['economy'] == economy].copy().iloc[:, 1:-2]

    netz_all_elecheat = historical_eh2.merge(netz_elecheat, how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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

    modren_prop1 = ['Modern renewables', 'Carbon neutrality'] + list(netz_modren_3.iloc[netz_modren_3.shape[0] - 4, 2:] / netz_modren_3.iloc[netz_modren_3.shape[0] - 1, 2:])
    modren_prop_series1 = pd.Series(modren_prop1, index = netz_modren_3.columns)

    netz_modren_4 = netz_modren_3.append([non_ren_series1, modren_prop_series1], ignore_index = True).reset_index(drop = True)

    # Remove historical from CN
    netz_modren_4.loc[netz_modren_4['item_code_new'] == 'Carbon neutrality', '2000':'2017'] = np.nan

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

        netz_ei_calc2 = [economy, 'Carbon neutrality'] + list(netz_enint_2.iloc[2, 2:] / netz_enint_2.iloc[2, 7] * 100)
        netz_ei_series2 = pd.Series(netz_ei_calc2, index = netz_enint_2.columns)

        netz_enint_3 = netz_enint_2.append(netz_ei_series2, ignore_index = True).reset_index(drop = True)

        netz_enint_3.loc[netz_enint_3['Series'] == 'Carbon neutrality', '2000':'2017'] = np.nan

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

        ref_calc1 = [economy, 'Reference'] + list(ref_enint_sup1.iloc[0, 2:] / ref_enint_sup1.iloc[1, 2:])
        ref_series1 = pd.Series(ref_calc1, index = ref_enint_sup1.columns)

        ref_enint_sup2 = ref_enint_sup1.append(ref_series1, ignore_index = True).reset_index(drop = True)

        ref_calc2 = [economy, 'Reference (indexed)'] + list(ref_enint_sup2.iloc[2, 2:] / ref_enint_sup2.iloc[2, 7] * 100)
        ref_series2 = pd.Series(ref_calc2, index = ref_enint_sup2.columns)

        ref_enint_sup3 = ref_enint_sup2.append(ref_series2, ignore_index = True).reset_index(drop = True)

        ref_enint_sup3_rows = ref_enint_sup3.shape[0]
        ref_enint_sup3_cols = ref_enint_sup3.shape[1]

        # CARBON NEUTRALITY
        netz_enint_sup1 = netz_tpes_1[netz_tpes_1['fuel_code'] == 'Total'].copy().reset_index(drop = True)
        netz_enint_sup1['Economy'] = economy
        netz_enint_sup1['Series'] = 'TPES'

        netz_enint_sup1 = netz_enint_sup1.append(macro_1[macro_1['Series'] == 'GDP 2018 USD PPP']).copy().reset_index(drop = True)

        netz_enint_sup1 = netz_enint_sup1[['Economy', 'Series'] + list(netz_enint_sup1.loc[:, '2000':'2050'])]

        netz_calc1 = [economy, 'Carbon neutrality'] + list(netz_enint_sup1.iloc[0, 2:] / netz_enint_sup1.iloc[1, 2:])
        netz_series1 = pd.Series(netz_calc1, index = netz_enint_sup1.columns)

        netz_enint_sup2 = netz_enint_sup1.append(netz_series1, ignore_index = True).reset_index(drop = True)

        netz_calc2 = [economy, 'Carbon neutrality (indexed)'] + list(netz_enint_sup2.iloc[2, 2:] / netz_enint_sup2.iloc[2, 7] * 100)
        netz_series2 = pd.Series(netz_calc2, index = netz_enint_sup2.columns)

        netz_enint_sup3 = netz_enint_sup2.append(netz_series2, ignore_index = True).reset_index(drop = True)

        # Remove CN historical
        netz_enint_sup3.loc[netz_enint_sup3['Series'] == 'Carbon neutrality', '2000':'2017'] = np.nan

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

    emiss_total_1.loc[emiss_total_1['fuel_code'] == 'Total', 'fuel_code'] = 'Carbon neutrality'

    # Remove historical from carbon neutrality
    emiss_total_1.loc[emiss_total_1['fuel_code'] == 'Carbon neutrality', '2000':'2017'] = np.nan

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
    emissions_wedge_1.loc[emissions_wedge_1['fuel_code'] == 'Carbon neutrality', 'item_code_new'] = 'Carbon neutrality'
    emissions_wedge_1.loc[emissions_wedge_1['item_code_new'] == 'Reference', 'fuel_code'] = '19_total'
    emissions_wedge_1.loc[emissions_wedge_1['item_code_new'] == 'Carbon neutrality', 'fuel_code'] = '19_total'

    # Get rid of data for CN for historical
    emissions_wedge_1.loc[emissions_wedge_1['item_code_new'] == 'Carbon neutrality', '2000':'2017'] = np.nan

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

    emissions_wedge_2.loc[emissions_wedge_2['fuel_code'] == 'Carbon neutrality', 'fuel_code'] = 'Carbon neutrality'

    # Get rid of data for CN for historical
    emissions_wedge_2.loc[emissions_wedge_2['fuel_code'] == 'Carbon neutrality', '2000':'2017'] = np.nan

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
                                        iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_hist_own_oil.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_crude_own = hist_ownoil.merge(ref_crude_own, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_crude_own = ref_crude_own[['fuel_code', 'item_code_new'] + list(ref_crude_own.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    # Power
    ref_crude_power = ref_power_df1[(ref_power_df1['economy'] == economy) &
                                    (ref_power_df1['FUEL'].isin(['6_1_crude_oil', '6_x_ngls']))].copy().reset_index(drop = True)

    ref_crude_power = ref_crude_power.groupby(['economy']).sum().copy().reset_index(drop = True)\
                          .assign(fuel_code = 'Crude oil & NGL', item_code_new = 'Power')

    #################################################################################
    hist_poweroil = EGEDA_histpower_oil[(EGEDA_histpower_oil['economy'] == economy) &
                                        (EGEDA_histpower_oil['FUEL'] == 'Crude oil & NGL')].copy()\
                                            .iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_histpower_oil.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_crude_power = hist_poweroil.merge(ref_crude_power, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_crude_power = ref_crude_power[['fuel_code', 'item_code_new'] + list(ref_crude_power.loc[:, '2000':'2050'])].copy().reset_index(drop = True)
    
    # Refining
    ref_crude_refinery = ref_refinery_1[ref_refinery_1['FUEL'].isin(['Crude oil', 'NGLs'])]\
                                .copy().groupby(['Transformation']).sum().reset_index(drop = True)\
                                    .assign(fuel_code = '6_crude_oil_and_ngl', item_code_new = '9_4_oil_refineries')

    hist_refine = EGEDA_hist_refining[EGEDA_hist_refining['economy'] == economy].copy()\
                     .iloc[:,:-2][['fuel_code', 'item_code_new'] + list(EGEDA_hist_refining.loc[:, '2000':'2016'])]\
                     .reset_index(drop = True)


    ref_crude_refinery = hist_refine.merge(ref_crude_refinery, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
                                        iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_hist_own_oil.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_petprod_own = hist_ownoil.merge(ref_petprod_own, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
                                            .iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_histpower_oil.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_petprod_power = hist_poweroil.merge(ref_petprod_power, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
                                                     [['fuel_code', 'item_code_new'] + list(EGEDA_hist_own_renew.loc[:, '2000':'2016'])]

    ref_renew_own = hist_ownrenew.merge(ref_renew_own, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
                                            .iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_histpower_renew.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_renew_power = hist_powerrenew.merge(ref_renew_power, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
                                        iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_hist_own_oil.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_crude_own = hist_ownoil.merge(netz_crude_own, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_crude_own = netz_crude_own[['fuel_code', 'item_code_new'] + list(netz_crude_own.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    # Power
    netz_crude_power = netz_power_df1[(netz_power_df1['economy'] == economy) &
                                    (netz_power_df1['FUEL'].isin(['6_1_crude_oil', '6_x_ngls']))].copy().reset_index(drop = True)

    netz_crude_power = netz_crude_power.groupby(['economy']).sum().copy().reset_index(drop = True)\
                          .assign(fuel_code = 'Crude oil & NGL', item_code_new = 'Power')

    #################################################################################
    hist_poweroil = EGEDA_histpower_oil[(EGEDA_histpower_oil['economy'] == economy) &
                                        (EGEDA_histpower_oil['FUEL'] == 'Crude oil & NGL')].copy()\
                                            .iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_histpower_oil.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_crude_power = hist_poweroil.merge(netz_crude_power, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_crude_power = netz_crude_power[['fuel_code', 'item_code_new'] + list(netz_crude_power.loc[:, '2000':'2050'])].copy().reset_index(drop = True)
    
    # Refining
    netz_crude_refinery = netz_refinery_1[netz_refinery_1['FUEL'].isin(['Crude oil', 'NGLs'])]\
                                .copy().groupby(['Transformation']).sum().reset_index(drop = True)\
                                    .assign(fuel_code = '6_crude_oil_and_ngl', item_code_new = '9_4_oil_refineries')

    hist_refine = EGEDA_hist_refining[EGEDA_hist_refining['economy'] == economy].copy()\
                     .iloc[:,:-2][['fuel_code', 'item_code_new'] + list(EGEDA_hist_refining.loc[:, '2000':'2016'])]\
                     .reset_index(drop = True)


    netz_crude_refinery = hist_refine.merge(netz_crude_refinery, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
                                        iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_hist_own_oil.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_petprod_own = hist_ownoil.merge(netz_petprod_own, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
                                            .iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_histpower_oil.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_petprod_power = hist_poweroil.merge(netz_petprod_power, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
                                                     [['fuel_code', 'item_code_new'] + list(EGEDA_hist_own_renew.loc[:, '2000':'2016'])]

    netz_renew_own = hist_ownrenew.merge(netz_renew_own, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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
                                            .iloc[:,:-2][['FUEL', 'item_code_new'] + list(EGEDA_histpower_renew.loc[:, '2000':'2016'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_renew_power = hist_powerrenew.merge(netz_renew_power, how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

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

    netz_calc1 = ['Carbon neutrality', 'CO2 intensity'] + list(netz_co2int_1.iloc[0, 2:] / netz_co2int_1.iloc[1, 2:])
    netz_series1 = pd.Series(netz_calc1, index = netz_co2int_1.columns)

    netz_co2int_2 = netz_co2int_1.append(netz_series1, ignore_index = True).reset_index(drop = True)

    # Remove 2000 to 2017 data from CN
    netz_co2int_2.loc[netz_co2int_2['fuel_code'] == 'Carbon neutrality', '2000':'2017'] = np.nan

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
        netz_ei_growth = (netz_enint_sup3.loc[netz_enint_sup3['Series'] == 'Carbon neutrality', '2050'] / netz_enint_sup3.loc[netz_enint_sup3['Series'] == 'Carbon neutrality', '2018']).to_numpy()
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
                                    columns = ['Carbon neutrality', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            netz_kaya_1.loc[0, 'Carbon neutrality'] = 'initial'
            netz_kaya_1.loc[1, 'Carbon neutrality'] = 'empty'
            netz_kaya_1.loc[2, 'Carbon neutrality'] = 'no improve'
            netz_kaya_1.loc[3, 'Carbon neutrality'] = 'empty'
            netz_kaya_1.loc[4, 'Carbon neutrality'] = 'no improve'
            netz_kaya_1.loc[5, 'Carbon neutrality'] = 'improve'
            netz_kaya_1.loc[6, 'Carbon neutrality'] = 'improve'

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
                                    columns = ['Carbon neutrality', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            netz_kaya_1.loc[0, 'Carbon neutrality'] = 'initial'
            netz_kaya_1.loc[1, 'Carbon neutrality'] = 'empty'
            netz_kaya_1.loc[2, 'Carbon neutrality'] = 'improve'
            netz_kaya_1.loc[3, 'Carbon neutrality'] = 'empty'
            netz_kaya_1.loc[4, 'Carbon neutrality'] = 'no improve'
            netz_kaya_1.loc[5, 'Carbon neutrality'] = 'improve'
            netz_kaya_1.loc[6, 'Carbon neutrality'] = 'improve'

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
                                    columns = ['Carbon neutrality', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            netz_kaya_1.loc[0, 'Carbon neutrality'] = 'initial'
            netz_kaya_1.loc[1, 'Carbon neutrality'] = 'empty'
            netz_kaya_1.loc[2, 'Carbon neutrality'] = 'no improve'
            netz_kaya_1.loc[3, 'Carbon neutrality'] = 'empty'
            netz_kaya_1.loc[4, 'Carbon neutrality'] = 'no improve'
            netz_kaya_1.loc[5, 'Carbon neutrality'] = 'improve'
            netz_kaya_1.loc[6, 'Carbon neutrality'] = 'no improve'

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
                                    columns = ['Carbon neutrality', 'Emissions 2018', 'Population', 'GDP per capita',\
                                                'Energy intensity', 'Emissions intensity', 'Emissions 2050'])

            netz_kaya_1.loc[0, 'Carbon neutrality'] = 'initial'
            netz_kaya_1.loc[1, 'Carbon neutrality'] = 'empty'
            netz_kaya_1.loc[2, 'Carbon neutrality'] = 'improve'
            netz_kaya_1.loc[3, 'Carbon neutrality'] = 'empty'
            netz_kaya_1.loc[4, 'Carbon neutrality'] = 'no improve'
            netz_kaya_1.loc[5, 'Carbon neutrality'] = 'improve'
            netz_kaya_1.loc[6, 'Carbon neutrality'] = 'no improve'

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

    ##############################################################################################################################
    
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

    # FED
    ref_fedfuel_1.to_excel(writer, sheet_name = 'FED by fuel', index = False, startrow = chart_height)
    netz_fedfuel_1.to_excel(writer, sheet_name = 'FED by fuel', index = False, startrow = (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6)
    ref_fedfuel_2.to_excel(writer, sheet_name = 'FED by fuel', index = False, startrow = chart_height + ref_fedfuel_1_rows + 3)
    netz_fedfuel_2.to_excel(writer, sheet_name = 'FED by fuel', index = False, startrow = (2 * chart_height) + ref_fedfuel_1_rows + netz_fedfuel_1_rows + ref_fedfuel_2_rows + 9)
    ref_fedsector_2.to_excel(writer, sheet_name = 'FED by sector', index = False, startrow = chart_height)
    netz_fedsector_2.to_excel(writer, sheet_name = 'FED by sector', index = False, startrow = (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9)
    ref_fedsector_3.to_excel(writer, sheet_name = 'FED by sector', index = False, startrow = chart_height + ref_fedsector_2_rows + 3)
    netz_fedsector_3.to_excel(writer, sheet_name = 'FED by sector', index = False, startrow = (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + 12)
    ref_tfec_1.to_excel(writer, sheet_name = 'FED by sector', index = False, startrow = chart_height + ref_fedsector_2_rows + ref_fedsector_3_rows + 6)
    netz_tfec_1.to_excel(writer, sheet_name = 'FED by sector', index = False, startrow = (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + netz_fedsector_3_rows + 15)
    ref_bld_2.to_excel(writer, sheet_name = 'Buildings', index = False, startrow = chart_height)
    netz_bld_2.to_excel(writer, sheet_name = 'Buildings', index = False, startrow = (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + 6)
    ref_bld_3.to_excel(writer, sheet_name = 'Buildings', index = False, startrow = chart_height + ref_bld_2_rows + 3)
    netz_bld_3.to_excel(writer, sheet_name = 'Buildings', index = False, startrow = (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 9)
    ref_ind_1.to_excel(writer, sheet_name = 'Industry', index = False, startrow = chart_height)
    netz_ind_1.to_excel(writer, sheet_name = 'Industry', index = False, startrow = (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + 6)
    ref_ind_2.to_excel(writer, sheet_name = 'Industry', index = False, startrow = chart_height + ref_ind_1_rows + 3)
    netz_ind_2.to_excel(writer, sheet_name = 'Industry', index = False, startrow = (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + 9)
    ref_steel_3.to_excel(writer, sheet_name = 'Heavy industry', index = False, startrow = chart_height)
    ref_chem_3.to_excel(writer, sheet_name = 'Heavy industry', index = False, startrow = chart_height + ref_steel_3_rows + 3)
    ref_cement_3.to_excel(writer, sheet_name = 'Heavy industry', index = False, startrow = chart_height + ref_steel_3_rows + ref_chem_3_rows + 6)
    netz_steel_3.to_excel(writer, sheet_name = 'Heavy industry', index = False, startrow = (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + 9)
    netz_chem_3.to_excel(writer, sheet_name = 'Heavy industry', index = False, startrow = (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + 12)
    netz_cement_3.to_excel(writer, sheet_name = 'Heavy industry', index = False, startrow = (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + netz_chem_3_rows + 15)
    ref_trn_1.to_excel(writer, sheet_name = 'Transport', index = False, startrow = chart_height)
    netz_trn_1.to_excel(writer, sheet_name = 'Transport', index = False, startrow = (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + 6)
    ref_trn_2.to_excel(writer, sheet_name = 'Transport', index = False, startrow = chart_height + ref_trn_1_rows + 3)
    netz_trn_2.to_excel(writer, sheet_name = 'Transport', index = False, startrow = (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + 9)
    ref_roadmodal_3.to_excel(writer, sheet_name = 'Road transport', index = False, startrow = chart_height)
    ref_roadfuel_3.to_excel(writer, sheet_name = 'Road transport', index = False, startrow = chart_height + ref_roadmodal_3_rows + 3)
    netz_roadmodal_3.to_excel(writer, sheet_name = 'Road transport', index = False, startrow = (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + 6)
    netz_roadfuel_3.to_excel(writer, sheet_name = 'Road transport', index = False, startrow = (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + netz_roadmodal_3_rows + 9)
    ref_ag_1.to_excel(writer, sheet_name = 'Agriculture', index = False, startrow = chart_height)
    netz_ag_1.to_excel(writer, sheet_name = 'Agriculture', index = False, startrow = (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6)
    ref_ag_2.to_excel(writer, sheet_name = 'Agriculture', index = False, startrow = chart_height + ref_ag_1_rows + 3)
    netz_ag_2.to_excel(writer, sheet_name = 'Agriculture', index = False, startrow = (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + 9)

    # Transformation
    ref_elecgen_2.to_excel(writer, sheet_name = 'Generation', index = False, startrow = chart_height)
    netz_elecgen_2.to_excel(writer, sheet_name = 'Generation', index = False, startrow = (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6)
    ref_elecgen_3.to_excel(writer, sheet_name = 'Generation', index = False, startrow = chart_height + ref_elecgen_2_rows + 3)
    netz_elecgen_3.to_excel(writer, sheet_name = 'Generation', index = False, startrow = (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9)
    ref_powcap_1.to_excel(writer, sheet_name = 'Capacity', index = False, startrow = chart_height)
    netz_powcap_1.to_excel(writer, sheet_name = 'Capacity', index = False, startrow = (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6)
    ref_powcap_2.to_excel(writer, sheet_name = 'Capacity', index = False, startrow = chart_height + ref_powcap_1_rows + 3)
    netz_powcap_2.to_excel(writer, sheet_name = 'Capacity', index = False, startrow = (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9)
    ref_pow_use_2.to_excel(writer, sheet_name = 'Power fuel consumption', index = False, startrow = chart_height)
    netz_pow_use_2.to_excel(writer, sheet_name = 'Power fuel consumption', index = False, startrow = (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + 6)
    ref_pow_use_3.to_excel(writer, sheet_name = 'Power fuel consumption', index = False, startrow = chart_height + ref_pow_use_2_rows + 3)
    netz_pow_use_3.to_excel(writer, sheet_name = 'Power fuel consumption', index = False, startrow = (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + 9)
    ref_refinery_1.to_excel(writer, sheet_name = 'Refining', index = False, startrow = chart_height)
    netz_refinery_1.to_excel(writer, sheet_name = 'Refining', index = False, startrow = (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9)
    ref_refinery_2.to_excel(writer, sheet_name = 'Refining', index = False, startrow = chart_height + ref_refinery_1_rows + 3)
    netz_refinery_2.to_excel(writer, sheet_name = 'Refining', index = False, startrow = (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + 12)
    ref_refinery_3.to_excel(writer, sheet_name = 'Refining', index = False, startrow = chart_height + ref_refinery_1_rows + ref_refinery_2_rows + 6)
    netz_refinery_3.to_excel(writer, sheet_name = 'Refining', index = False, startrow = (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + 15)
    ref_trans_3.to_excel(writer, sheet_name = 'Transformation', index = False, startrow = chart_height)
    netz_trans_3.to_excel(writer, sheet_name = 'Transformation', index = False, startrow = (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6)
    ref_trans_4.to_excel(writer, sheet_name = 'Transformation', index = False, startrow = chart_height + ref_trans_3_rows + 3)
    netz_trans_4.to_excel(writer, sheet_name = 'Transformation', index = False, startrow = (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + 9)
    ref_ownuse_1.to_excel(writer, sheet_name = 'Own-use', index = False, startrow = chart_height)
    netz_ownuse_1.to_excel(writer, sheet_name = 'Own-use', index = False, startrow = (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + 6)
    ref_ownuse_2.to_excel(writer, sheet_name = 'Own-use', index = False, startrow = chart_height + ref_ownuse_1_rows + 3)
    netz_ownuse_2.to_excel(writer, sheet_name = 'Own-use', index = False, startrow = (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + 9)
    ref_heat_use_2.to_excel(writer, sheet_name = 'Heat only consumption', index = False, startrow = chart_height)
    netz_heat_use_2.to_excel(writer, sheet_name = 'Heat only consumption', index = False, startrow = (2 * chart_height) + ref_heat_use_2_rows + ref_heat_use_3_rows + 6)
    ref_heat_use_3.to_excel(writer, sheet_name = 'Heat only consumption', index = False, startrow = chart_height + ref_heat_use_2_rows + 3)
    netz_heat_use_3.to_excel(writer, sheet_name = 'Heat only consumption', index = False, startrow = (2 * chart_height) + ref_heat_use_2_rows + ref_heat_use_3_rows + netz_heat_use_2_rows + 9)
    ref_heatgen_2.to_excel(writer, sheet_name = 'Heat generation', index = False, startrow = chart_height)
    netz_heatgen_2.to_excel(writer, sheet_name = 'Heat generation', index = False, startrow = (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6)
    ref_heatgen_3.to_excel(writer, sheet_name = 'Heat generation', index = False, startrow = chart_height + ref_heatgen_2_rows + 3)
    netz_heatgen_3.to_excel(writer, sheet_name = 'Heat generation', index = False, startrow = (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9)

    # TPES
    ref_tpes_1.to_excel(writer, sheet_name = 'Supply', index = False, startrow = chart_height)
    netz_tpes_1.to_excel(writer, sheet_name = 'Supply', index = False, startrow = (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6)
    ref_tpes_2.to_excel(writer, sheet_name = 'Supply', index = False, startrow = chart_height + ref_tpes_1_rows + 3)
    netz_tpes_2.to_excel(writer, sheet_name = 'Supply', index = False, startrow = (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + 9)
    ref_prod_1.to_excel(writer, sheet_name = 'Production', index = False, startrow = chart_height)
    netz_prod_1.to_excel(writer, sheet_name = 'Production', index = False, startrow = (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6)
    ref_prod_2.to_excel(writer, sheet_name = 'Production', index = False, startrow = chart_height + ref_prod_1_rows + 3)
    netz_prod_2.to_excel(writer, sheet_name = 'Production', index = False, startrow = (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + 9)
    ref_tpes_comp_1.to_excel(writer, sheet_name = 'Supply REF', index = False, startrow = chart_height)
    netz_tpes_comp_1.to_excel(writer, sheet_name = 'Supply CN', index = False, startrow = chart_height)
    ref_imports_1.to_excel(writer, sheet_name = 'Supply REF', index = False, startrow = chart_height + ref_tpes_comp_1_rows + 3)
    netz_imports_1.to_excel(writer, sheet_name = 'Supply CN', index = False, startrow = chart_height + netz_tpes_comp_1_rows + 3)
    ref_imports_2.to_excel(writer, sheet_name = 'Supply REF', index = False, startrow = chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + 6)
    netz_imports_2.to_excel(writer, sheet_name = 'Supply CN', index = False, startrow = chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + 6)
    ref_exports_1.to_excel(writer, sheet_name = 'Supply REF', index = False, startrow = chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + 9)
    netz_exports_1.to_excel(writer, sheet_name = 'Supply CN', index = False, startrow = chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + 9)
    ref_exports_2.to_excel(writer, sheet_name = 'Supply REF', index = False, startrow = chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + 12)
    ref_electrade_1.to_excel(writer, sheet_name = 'Supply REF', index = False, startrow = chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + 15)
    ref_nettrade_1.to_excel(writer, sheet_name = 'Supply REF', index = False, startrow = chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + ref_electrade_1_rows + 18)   
    netz_exports_2.to_excel(writer, sheet_name = 'Supply CN', index = False, startrow = chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + 12)
    netz_electrade_1.to_excel(writer, sheet_name = 'Supply CN', index = False, startrow = chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + 15)
    netz_nettrade_1.to_excel(writer, sheet_name = 'Supply CN', index = False, startrow = chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + netz_electrade_1_rows + 18)
    ref_bunkers_1.to_excel(writer, sheet_name = 'Bunkers', index = False, startrow = chart_height)
    netz_bunkers_1.to_excel(writer, sheet_name = 'Bunkers', index = False, startrow = (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + 6)
    ref_bunkers_2.to_excel(writer, sheet_name = 'Bunkers', index = False, startrow = chart_height + ref_bunkers_1_rows + 3)
    netz_bunkers_2.to_excel(writer, sheet_name = 'Bunkers', index = False, startrow = (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + 9)

    # Fuels
    ref_coalcons_1.to_excel(writer, sheet_name = 'Coal', index = False, startrow = chart_height)
    ref_coal_1.to_excel(writer, sheet_name = 'Coal', index = False, startrow = chart_height + ref_coalcons_1_rows + 3)
    ref_ct_prod1.to_excel(writer, sheet_name = 'Coal by type', index = False, startrow = chart_height)
    ref_ct_imports1.to_excel(writer, sheet_name = 'Coal by type', index = False, startrow = chart_height + ref_ct_prod1_rows + 3)
    ref_ct_exports1.to_excel(writer, sheet_name = 'Coal by type', index = False, startrow = chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + 6)
    ref_gascons_1.to_excel(writer, sheet_name = 'Gas', index = False, startrow = chart_height)
    ref_gas_1.to_excel(writer, sheet_name = 'Gas', index = False, startrow = chart_height + ref_gascons_1_rows + 3)
    ref_gasim_1.to_excel(writer, sheet_name = 'Gas trade', index = False, startrow = chart_height)
    ref_gasex_1.to_excel(writer, sheet_name = 'Gas trade', index = False, startrow = chart_height + ref_gasim_1_rows + 3)
    ref_crudecons_1.to_excel(writer, sheet_name = 'Crude & NGL', index = False, startrow = chart_height)
    ref_crude_1.to_excel(writer, sheet_name = 'Crude & NGL', index = False, startrow = chart_height + ref_crudecons_1_rows + 3)
    ref_petprodcons_1.to_excel(writer, sheet_name = 'Petroleum products', index = False, startrow = chart_height)
    ref_petprod_2.to_excel(writer, sheet_name = 'Petroleum products', index = False, startrow = chart_height + ref_petprodcons_1_rows + 3)
    netz_coalcons_1.to_excel(writer, sheet_name = 'Coal', index = False, startrow = (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + 6)
    netz_coal_1.to_excel(writer, sheet_name = 'Coal', index = False, startrow = (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + netz_coalcons_1_rows + 9)
    netz_ct_prod1.to_excel(writer, sheet_name = 'Coal by type', index = False, startrow = (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + 9)
    netz_ct_imports1.to_excel(writer, sheet_name = 'Coal by type', index = False, startrow = (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + 12)
    netz_ct_exports1.to_excel(writer, sheet_name = 'Coal by type', index = False, startrow = (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + netz_ct_imports1_rows + 15)
    ref_coaltpes_2.to_excel(writer, sheet_name = 'Coal by type', index = False, startrow = (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + netz_ct_imports1_rows + netz_ct_exports1_rows + 21)
    netz_coaltpes_2.to_excel(writer, sheet_name = 'Coal by type', index = False, startrow = (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + netz_ct_imports1_rows + netz_ct_exports1_rows + ref_coaltpes_2_rows + 24)
    netz_gascons_1.to_excel(writer, sheet_name = 'Gas', index = False, startrow = (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + 6)
    netz_gas_1.to_excel(writer, sheet_name = 'Gas', index = False, startrow = (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + netz_gascons_1_rows + 9)
    netz_gasim_1.to_excel(writer, sheet_name = 'Gas trade', index = False, startrow = (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + 6)
    netz_gasex_1.to_excel(writer, sheet_name = 'Gas trade', index = False, startrow = (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + netz_gasim_1_rows + 9)
    netz_crudecons_1.to_excel(writer, sheet_name = 'Crude & NGL', index = False, startrow = (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + 6)
    netz_crude_1.to_excel(writer, sheet_name = 'Crude & NGL', index = False, startrow = (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + netz_crudecons_1_rows + 9)
    netz_petprodcons_1.to_excel(writer, sheet_name = 'Petroleum products', index = False, startrow = (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + 6)
    netz_petprod_2.to_excel(writer, sheet_name = 'Petroleum products', index = False, startrow = (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + netz_petprodcons_1_rows + 9)
    ref_renewcons_1.to_excel(writer, sheet_name = 'Renewable fuels', index = False, startrow = chart_height)
    ref_renew_2.to_excel(writer, sheet_name = 'Renewable fuels', index = False, startrow = chart_height + ref_renewcons_1_rows + 3)
    netz_renewcons_1.to_excel(writer, sheet_name = 'Renewable fuels', index = False, startrow = (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + 6)
    netz_renew_2.to_excel(writer, sheet_name = 'Renewable fuels', index = False, startrow = (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + netz_renewcons_1_rows + 9)
    ref_renewcons_2.to_excel(writer, sheet_name = 'Renewable fuels VER2', index = False, startrow = chart_height)
    ref_renew_2.to_excel(writer, sheet_name = 'Renewable fuels VER2', index = False, startrow = chart_height + ref_renewcons_2_rows + 3)
    netz_renewcons_2.to_excel(writer, sheet_name = 'Renewable fuels VER2', index = False, startrow = (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + 6)
    netz_renew_2.to_excel(writer, sheet_name = 'Renewable fuels VER2', index = False, startrow = (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + netz_renewcons_2_rows + 9)
    ref_elec_1.to_excel(writer, sheet_name = 'Electricity', index = False, startrow = chart_height)
    netz_elec_1.to_excel(writer, sheet_name = 'Electricity', index = False, startrow = (2 * chart_height) + ref_elec_1_rows + 3)
    ref_hyd_1.to_excel(writer, sheet_name = 'Hydrogen', index = False, startrow = chart_height)
    ref_hydrogen_3.to_excel(writer, sheet_name = 'Hydrogen', index = False, startrow = chart_height + ref_hyd_1_rows + 3)
    ref_hyd_use_1.to_excel(writer, sheet_name = 'Hydrogen', index = False, startrow = chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + 6)
    netz_hyd_1.to_excel(writer, sheet_name = 'Hydrogen', index = False, startrow = (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + 9)
    netz_hydrogen_3.to_excel(writer, sheet_name = 'Hydrogen', index = False, startrow = (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + netz_hyd_1_rows + 12)
    netz_hyd_use_1.to_excel(writer, sheet_name = 'Hydrogen', index = False, startrow = (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + netz_hyd_1_rows + netz_hydrogen_3_rows + 15)

    # More fuels
    ref_nuke_1.to_excel(writer, sheet_name = 'Other fuels REF', index = False, startrow = chart_height)
    ref_biomass_1.to_excel(writer, sheet_name = 'Other fuels REF', index = False, startrow = chart_height + ref_nuke_1_rows + 3)
    ref_biofuel_2.to_excel(writer, sheet_name = 'Other fuels REF', index = False, startrow = chart_height + ref_nuke_1_rows + ref_biomass_1_rows + 6)
    netz_nuke_1.to_excel(writer, sheet_name = 'Other fuels CN', index = False, startrow = chart_height)
    netz_biomass_1.to_excel(writer, sheet_name = 'Other fuels CN', index = False, startrow = chart_height + netz_nuke_1_rows + 3)
    netz_biofuel_2.to_excel(writer, sheet_name = 'Other fuels CN', index = False, startrow = chart_height + netz_nuke_1_rows + netz_biomass_1_rows + 6)

    # Miscellaneous 
    ref_modren_4.to_excel(writer, sheet_name = 'Modern renewables', index = False, startrow = chart_height)
    netz_modren_4.to_excel(writer, sheet_name = 'Modern renewables', index = False, startrow = chart_height + ref_modren_4_rows + 3)
    ref_enint_3.to_excel(writer, sheet_name = 'Energy intensity', index = False, startrow = chart_height)
    netz_enint_3.to_excel(writer, sheet_name = 'Energy intensity', index = False, startrow = chart_height + ref_enint_3_rows + 3)
    ref_enint_sup3.to_excel(writer, sheet_name = 'Energy intensity', index = False, startrow = chart_height + ref_enint_3_rows + netz_enint_3_rows + 6)
    netz_enint_sup3.to_excel(writer, sheet_name = 'Energy intensity', index = False, startrow = chart_height + ref_enint_3_rows + netz_enint_3_rows + ref_enint_sup3_rows + 9)
    ref_emiss_fuel_1.to_excel(writer, sheet_name = 'CO2 by fuel', index = False, startrow = chart_height)
    netz_emiss_fuel_1.to_excel(writer, sheet_name = 'CO2 by fuel', index = False, startrow = (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6)
    ref_emiss_fuel_2.to_excel(writer, sheet_name = 'CO2 by fuel', index = False, startrow = chart_height + ref_emiss_fuel_1_rows + 3)
    netz_emiss_fuel_2.to_excel(writer, sheet_name = 'CO2 by fuel', index = False, startrow = (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + 9)
    ref_emiss_sector_1.to_excel(writer, sheet_name = 'CO2 by sector', index = False, startrow = chart_height)
    netz_emiss_sector_1.to_excel(writer, sheet_name = 'CO2 by sector', index = False, startrow = (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6)
    ref_emiss_sector_2.to_excel(writer, sheet_name = 'CO2 by sector', index = False, startrow = chart_height + ref_emiss_sector_1_rows + 3)
    netz_emiss_sector_2.to_excel(writer, sheet_name = 'CO2 by sector', index = False, startrow = (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + 9)
    emissions_wedge_1.to_excel(writer, sheet_name = 'CO2 wedge', index = False, startrow = chart_height)
    emissions_wedge_2.to_excel(writer, sheet_name = 'CO2 wedge', index = False, startrow = chart_height + emissions_wedge_1_rows + 3)
    ref_co2int_2.to_excel(writer, sheet_name = 'CO2 intensity', index = False, startrow = chart_height)
    netz_co2int_2.to_excel(writer, sheet_name = 'CO2 intensity', index = False, startrow = chart_height + ref_co2int_2_rows + 3)
    emiss_total_1.to_excel(writer, sheet_name = 'CO2 intensity', index = False, startrow = chart_height + ref_co2int_2_rows + netz_co2int_2_rows + 6)
    ref_kaya_1.to_excel(writer, sheet_name = 'CO2 breakdown', index = False, startrow = chart_height)
    netz_kaya_1.to_excel(writer, sheet_name = 'CO2 breakdown', index = False, startrow = chart_height + ref_kaya_1_rows + 3)

    ################################################################################################################################

    # CHARTS
    # REFERENCE

    # Access the workbook and first sheet with data from df1
    ref_worksheet1 = writer.sheets['FED by fuel']
    
    # Comma format and header format        
    space_format = workbook.add_format({'num_format': '# ### ### ##0.0;-# ### ### ##0.0;-'})
    percentage_format = workbook.add_format({'num_format': '0.0%'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
    cell_format2 = workbook.add_format({'font_size': 9})

        
    # Apply comma format and header format to relevant data rows
    ref_worksheet1.set_column(1, ref_fedfuel_1_cols + 1, None, space_format)
    ref_worksheet1.set_row(chart_height, None, header_format)
    ref_worksheet1.set_row(chart_height + ref_fedfuel_1_rows + 3, None, header_format)
    ref_worksheet1.set_row((2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, None, header_format)
    ref_worksheet1.set_row((2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + 9, None, header_format)
    ref_worksheet1.write(0, 0, economy + ' FED fuel Reference', cell_format1)
    ref_worksheet1.write(chart_height + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, 0, economy + ' FED fuel Carbon Neutrality', cell_format1)
    ref_worksheet1.write(1, 0, 'Units: Petajoules', cell_format2)

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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedfuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_fedfuel_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fedfuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_fedfuel_1_rows):
        if not ref_fedfuel_1['fuel_code'].iloc[i] in ['Total']:
            ref_fedfuel_chart1.add_series({
                'name':       ['FED by fuel', chart_height + i + 1, 0],
                'categories': ['FED by fuel', chart_height, 2, chart_height, ref_fedfuel_1_cols - 1],
                'values':     ['FED by fuel', chart_height + i + 1, 2, chart_height + i + 1, ref_fedfuel_1_cols - 1],
                'fill':       {'color': ref_fedfuel_1['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
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
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedfuel_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_fedfuel_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fedfuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in ref_fedfuel_2['fuel_code'].unique():
        i = ref_fedfuel_2[ref_fedfuel_2['fuel_code'] == component].index[0]
        if not ref_fedfuel_2['fuel_code'].iloc[i] in ['Total']:
            ref_fedfuel_chart2.add_series({
                'name':       ['FED by fuel', chart_height + ref_fedfuel_1_rows + i + 4, 0],
                'categories': ['FED by fuel', chart_height + ref_fedfuel_1_rows + 3, 2, chart_height + ref_fedfuel_1_rows + 3, ref_fedfuel_2_cols - 1],
                'values':     ['FED by fuel', chart_height + ref_fedfuel_1_rows + i + 4, 2, chart_height + ref_fedfuel_1_rows + i + 4, ref_fedfuel_2_cols - 1],
                'fill':       {'color': ref_fedfuel_2['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })

        else:
            pass
    
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedfuel_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_fedfuel_chart3.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fedfuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_fedfuel_1_rows):
        if not ref_fedfuel_1['fuel_code'].iloc[i] in ['Total']:
            ref_fedfuel_chart3.add_series({
                'name':       ['FED by fuel', chart_height + i + 1, 0],
                'categories': ['FED by fuel', chart_height, 2, chart_height, ref_fedfuel_1_cols - 1],
                'values':     ['FED by fuel', chart_height + i + 1, 2, chart_height + i + 1, ref_fedfuel_1_cols - 1],
                'line':       {'color': ref_fedfuel_1['fuel_code'].map(colours_dict).loc[i], 'width': 1}
            })

        else:
            pass    
        
    ref_worksheet1.insert_chart('R3', ref_fedfuel_chart3)

    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    ref_worksheet2 = writer.sheets['FED by sector']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet2.set_column(1, ref_fedsector_2_cols + 1, None, space_format)
    ref_worksheet2.set_row(chart_height, None, header_format)
    ref_worksheet2.set_row(chart_height + ref_fedsector_2_rows + 3, None, header_format)
    ref_worksheet2.set_row(chart_height + ref_fedsector_2_rows + ref_fedsector_3_rows + 6, None, header_format)
    ref_worksheet2.set_row((2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, None, header_format)
    ref_worksheet2.set_row((2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + 12, None, header_format)
    ref_worksheet2.set_row((2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + netz_fedsector_3_rows + 15, None, header_format)
    ref_worksheet2.write(0, 0, economy + ' FED sector Reference', cell_format1)
    ref_worksheet2.write(chart_height + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, 0, economy + ' FED sector Carbon Neutrality', cell_format1)
    ref_worksheet2.write(1, 0, 'Units: Petajoules', cell_format2)

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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedsector_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_fedsector_chart3.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fedsector_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_fedsector_2_rows):
        if not ref_fedsector_2['item_code_new'].iloc[i] in ['Total']:
            ref_fedsector_chart3.add_series({
                'name':       ['FED by sector', chart_height + i + 1, 1],
                'categories': ['FED by sector', chart_height, 2, chart_height, ref_fedsector_2_cols - 1],
                'values':     ['FED by sector', chart_height + i + 1, 2, chart_height + i + 1, ref_fedsector_2_cols - 1],
                'fill':       {'color': ref_fedsector_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
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
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedsector_chart4.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_fedsector_chart4.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fedsector_chart4.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in ref_fedsector_3['item_code_new'].unique():
        i = ref_fedsector_3[ref_fedsector_3['item_code_new'] == component].index[0]
        if not ref_fedsector_3['item_code_new'].iloc[i] in ['Total']:
            ref_fedsector_chart4.add_series({
                'name':       ['FED by sector', chart_height + ref_fedsector_2_rows + i + 4, 1],
                'categories': ['FED by sector', chart_height + ref_fedsector_2_rows + 3, 2, chart_height + ref_fedsector_2_rows + 3, ref_fedsector_3_cols - 1],
                'values':     ['FED by sector', chart_height + ref_fedsector_2_rows + i + 4, 2, chart_height + ref_fedsector_2_rows + i + 4, ref_fedsector_3_cols - 1],
                'fill':       {'color': ref_fedsector_3['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })

        else:
            pass
    
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedsector_chart5.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_fedsector_chart5.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fedsector_chart5.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_fedsector_2_rows):
        if not ref_fedsector_2['item_code_new'].iloc[i] in ['Total']:
            ref_fedsector_chart5.add_series({
                'name':       ['FED by sector', chart_height + i + 1, 1],
                'categories': ['FED by sector', chart_height, 2, chart_height, ref_fedsector_2_cols - 1],
                'values':     ['FED by sector', chart_height + i + 1, 2, chart_height + i + 1, ref_fedsector_2_cols - 1],
                'line':       {'color': ref_fedsector_2['item_code_new'].map(colours_dict).loc[i], 'width': 1}
            })

        else:
            pass    
        
    ref_worksheet2.insert_chart('R3', ref_fedsector_chart5)
    
    ############################# Next sheet: FED (TFC) for building sector ##################################
    
    # Access the workbook and third sheet with data from bld_df1
    ref_worksheet3 = writer.sheets['Buildings']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet3.set_column(2, ref_bld_2_cols + 1, None, space_format)
    ref_worksheet3.set_row(chart_height, None, header_format)
    ref_worksheet3.set_row(chart_height + ref_bld_2_rows + 3, None, header_format)
    ref_worksheet3.set_row((2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + 6, None, header_format)
    ref_worksheet3.set_row((2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 9, None, header_format)
    ref_worksheet3.write(0, 0, economy + ' buildings Reference', cell_format1)
    ref_worksheet3.write(chart_height + ref_bld_2_rows + ref_bld_3_rows + 6, 0, economy + ' buildings Carbon Neutrality', cell_format1)
    ref_worksheet3.write(1, 0, 'Units: Petajoules', cell_format2)

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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_fed_bld_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_fed_bld_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fed_bld_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ref_bld_2['fuel_code'].unique():
        i = ref_bld_2[ref_bld_2['fuel_code'] == component].index[0]
        if not ref_bld_2['fuel_code'].iloc[i] in ['Total']:
            ref_fed_bld_chart1.add_series({
                'name':       ['Buildings', chart_height + i + 1, 0],
                'categories': ['Buildings', chart_height, 2, chart_height, ref_bld_2_cols - 1],
                'values':     ['Buildings', chart_height + i + 1, 2, chart_height + i + 1, ref_bld_2_cols - 1],
                'fill':       {'color': ref_bld_2['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass

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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_fed_bld_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_fed_bld_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fed_bld_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for bld_sect in ['Services', 'Residential']:
        i = ref_bld_3[ref_bld_3['item_code_new'] == bld_sect].index[0]
        ref_fed_bld_chart2.add_series({
            'name':       ['Buildings', chart_height + ref_bld_2_rows + 4 + i, 1],
            'categories': ['Buildings', chart_height + ref_bld_2_rows + 3, 2, chart_height + ref_bld_2_rows + 3, ref_bld_3_cols - 1],
            'values':     ['Buildings', chart_height + ref_bld_2_rows + 4 + i, 2, chart_height + ref_bld_2_rows + 4 + i, ref_bld_3_cols - 1],
            'fill':       {'color': ref_bld_3['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet3.insert_chart('J3', ref_fed_bld_chart2)
    
    ############################# Next sheet: FED (TFC) for industry ##################################
    
    # Access the workbook and fourth sheet with data from bld_df1
    ref_worksheet4 = writer.sheets['Industry']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet4.set_column(2, ref_ind_1_cols + 1, None, space_format)
    ref_worksheet4.set_row(chart_height, None, header_format)
    ref_worksheet4.set_row(chart_height + ref_ind_1_rows + 3, None, header_format)
    ref_worksheet4.set_row((2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + 6, None, header_format)
    ref_worksheet4.set_row((2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + 9, None, header_format)
    ref_worksheet4.write(0, 0, economy + ' industry Reference', cell_format1)
    ref_worksheet4.write(chart_height + ref_ind_1_rows + ref_ind_2_rows + 6, 0, economy + ' industry Carbon Neutrality', cell_format1)
    ref_worksheet4.write(1, 0, 'Units: Petajoules', cell_format2)

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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_fed_ind_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_fed_ind_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fed_ind_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_ind_1_rows):
        if not ref_ind_1['item_code_new'].iloc[i] in ['Industry']:
            ref_fed_ind_chart1.add_series({
                'name':       ['Industry', chart_height + i + 1, 1],
                'categories': ['Industry', chart_height, 2, chart_height, ref_ind_1_cols - 1],
                'values':     ['Industry', chart_height + i + 1, 2, chart_height + i + 1, ref_ind_1_cols - 1],
                'fill':       {'color': ref_ind_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_fed_ind_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_fed_ind_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_fed_ind_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for fuel_agg in ref_ind_2['fuel_code'].unique():
        j = ref_ind_2[ref_ind_2['fuel_code'] == fuel_agg].index[0]
        if not ref_ind_2['fuel_code'].iloc[j] in ['Total']:
            ref_fed_ind_chart2.add_series({
                'name':       ['Industry', chart_height + ref_ind_1_rows + j + 4, 0],
                'categories': ['Industry', chart_height + ref_ind_1_rows + 3, 2, chart_height + ref_ind_1_rows + 3, ref_ind_2_cols - 1],
                'values':     ['Industry', chart_height + ref_ind_1_rows + j + 4, 2, chart_height + ref_ind_1_rows + j + 4, ref_ind_2_cols - 1],
                'fill':       {'color': ref_ind_2['fuel_code'].map(colours_dict).loc[j]},
                'border':     {'none': True}
            })

        else:
            pass
    
    ref_worksheet4.insert_chart('J3', ref_fed_ind_chart2)

    ################################# NEXT SHEET: TRANSPORT FED ################################################################

    # Access the workbook and first sheet with data from df1
    ref_worksheet5 = writer.sheets['Transport']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet5.set_column(2, ref_trn_1_cols + 1, None, space_format)
    ref_worksheet5.set_row(chart_height, None, header_format)
    ref_worksheet5.set_row(chart_height + ref_trn_1_rows + 3, None, header_format)
    ref_worksheet5.set_row((2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + 6, None, header_format)
    ref_worksheet5.set_row((2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + 9, None, header_format)
    ref_worksheet5.write(0, 0, economy + ' FED transport Reference', cell_format1)
    ref_worksheet5.write(chart_height + ref_trn_1_rows + ref_trn_2_rows + 6, 0, economy + ' FED transport Carbon Neutrality', cell_format1)
    ref_worksheet5.write(1, 0, 'Units: Petajoules', cell_format2)

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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_transport_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_transport_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_transport_chart1.set_title({
        'none': True
    })
        
    for fuel_agg in ref_trn_1['fuel_code'].unique():
        j = ref_trn_1[ref_trn_1['fuel_code'] == fuel_agg].index[0]
        if not ref_trn_1['fuel_code'].iloc[j] in ['Total']:
            ref_transport_chart1.add_series({
                'name':       ['Transport', chart_height + j + 1, 0],
                'categories': ['Transport', chart_height, 2, chart_height, ref_trn_1_cols - 1],
                'values':     ['Transport', chart_height + j + 1, 2, chart_height + j + 1, ref_trn_1_cols - 1],
                'fill':       {'color': ref_trn_1['fuel_code'].map(colours_dict).loc[j]},
                'border':     {'none': True} 
            })

        else:
            pass
    
    ref_worksheet5.insert_chart('B3', ref_transport_chart1)
            
    ############# FED transport chart 2 (transport by modality)
    
    # Create a FED transport modality column chart
    if ref_trn_2_rows > 0:
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_transport_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_transport_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_transport_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for modality in ref_trn_2['item_code_new'].unique():
            j = ref_trn_2[ref_trn_2['item_code_new'] == modality].index[0]
            if not ref_trn_2['item_code_new'].iloc[j] in ['Total']:
                ref_transport_chart2.add_series({
                    'name':       ['Transport', chart_height + ref_trn_1_rows + j + 4, 1],
                    'categories': ['Transport', chart_height + ref_trn_1_rows + 3, 2, chart_height + ref_trn_1_rows + 3, ref_trn_2_cols - 1],
                    'values':     ['Transport', chart_height + ref_trn_1_rows + j + 4, 2, chart_height + ref_trn_1_rows + j + 4, ref_trn_2_cols - 1],
                    'fill':       {'color': ref_trn_2['item_code_new'].map(colours_dict).loc[j]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        ref_worksheet5.insert_chart('J3', ref_transport_chart2)

    else:
        pass

    ################################# NEXT SHEET: AGRICULTURE FED ################################################################

    # Access the workbook and first sheet with data from df1
    ref_worksheet6 = writer.sheets['Agriculture']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet6.set_column(2, ref_ag_1_cols + 1, None, space_format)
    ref_worksheet6.set_row(chart_height, None, header_format)
    ref_worksheet6.set_row(chart_height + ref_ag_1_rows + 3, None, header_format)
    ref_worksheet6.set_row((2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, None, header_format)
    ref_worksheet6.set_row((2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + 9, None, header_format)
    ref_worksheet6.write(0, 0, economy + ' FED agriculture Reference', cell_format1)
    ref_worksheet6.write(chart_height + ref_ag_1_rows + ref_ag_2_rows + 6, 0, economy + ' FED agriculture Carbon Neutrality', cell_format1)
    ref_worksheet6.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a Agriculture line chart 
    if ref_ag_1_rows > 0:
        ref_ag_chart1 = workbook.add_chart({'type': 'line'})
        ref_ag_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_ag_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_ag_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_ag_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                    'width': 1,
                    'dash_type': 'square_dot'}
        })
            
        ref_ag_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_ag_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_ag_1_rows):
            if not ref_ag_1['fuel_code'].iloc[i] in ['Total']:
                ref_ag_chart1.add_series({
                    'name':       ['Agriculture', chart_height + i + 1, 0],
                    'categories': ['Agriculture', chart_height, 2, chart_height, ref_ag_1_cols - 1],
                    'values':     ['Agriculture', chart_height + i + 1, 2, chart_height + i + 1, ref_ag_1_cols - 1],
                    'line':       {'color': ref_ag_1['fuel_code'].map(colours_dict).loc[i], 'width': 1}
                })

            else:
                pass    
            
        ref_worksheet6.insert_chart('B3', ref_ag_chart1)

    else:
        pass

    # Create a Agriculture area chart
    if ref_ag_1_rows > 0:
        ref_ag_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_ag_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_ag_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        ref_ag_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_ag_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_ag_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_ag_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_ag_1_rows):
            if not ref_ag_1['fuel_code'].iloc[i] in ['Total']:
                ref_ag_chart2.add_series({
                    'name':       ['Agriculture', chart_height + i + 1, 0],
                    'categories': ['Agriculture', chart_height, 2, chart_height, ref_ag_1_cols - 1],
                    'values':     ['Agriculture', chart_height + i + 1, 2, chart_height + i + 1, ref_ag_1_cols - 1],
                    'fill':       {'color': ref_ag_1['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet6.insert_chart('J3', ref_ag_chart2)

    else:
        pass

    # Create a Agriculture stacked column
    if ref_ag_2_rows > 0:
        ref_ag_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
        ref_ag_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_ag_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        ref_ag_chart3.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'interval_unit': 1,
            'line': {'color': '#bebebe'}
        })
            
        ref_ag_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_ag_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_ag_chart3.set_title({
            'none': True
        })

        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_ag_2_rows):
            if not ref_ag_2['fuel_code'].iloc[i] in ['Total']:
                ref_ag_chart3.add_series({
                    'name':       ['Agriculture', chart_height + ref_ag_1_rows + i + 4, 0],
                    'categories': ['Agriculture', chart_height + ref_ag_1_rows + 3, 2, chart_height + ref_ag_1_rows + 3, ref_ag_2_cols - 1],
                    'values':     ['Agriculture', chart_height + ref_ag_1_rows + i + 4, 2, chart_height + ref_ag_1_rows + i + 4, ref_ag_2_cols - 1],
                    'fill':       {'color': ref_ag_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })
            
            else:
                pass
        
        ref_worksheet6.insert_chart('R3', ref_ag_chart3)

    else:
        pass

    # HYDROGEN CHARTS (REDUNDANT)

    
    ##############################################################################################################################

    # CHARTS
    # NET ZERO

    # Access the workbook and first sheet with data from df1
    # netz_worksheet1 = writer.sheets['FED by fuel']
    
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedfuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_fedfuel_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fedfuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_fedfuel_1_rows):
        if not netz_fedfuel_1['fuel_code'].iloc[i] in ['Total']:
            netz_fedfuel_chart1.add_series({
                'name':       ['FED by fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, 0],
                'categories': ['FED by fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, 2,\
                    (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, netz_fedfuel_1_cols - 1],
                'values':     ['FED by fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, netz_fedfuel_1_cols - 1],
                'fill':       {'color': netz_fedfuel_1['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
    ref_worksheet1.insert_chart('B' + str(chart_height + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 9), netz_fedfuel_chart1)

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
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedfuel_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_fedfuel_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fedfuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in netz_fedfuel_2['fuel_code'].unique():
        i = netz_fedfuel_2[netz_fedfuel_2['fuel_code'] == component].index[0]
        if not netz_fedfuel_2['fuel_code'].iloc[i] in ['Total']:
            netz_fedfuel_chart2.add_series({
                'name':       ['FED by fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + i + 10, 0],
                'categories': ['FED by fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + 9,\
                    2, (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + 9, netz_fedfuel_2_cols - 1],
                'values':     ['FED by fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + i + 10,\
                    2, (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + i + 10, netz_fedfuel_2_cols - 1],
                'fill':       {'color': netz_fedfuel_2['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })

        else:
            pass
    
    ref_worksheet1.insert_chart('J' + str(chart_height + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 9), netz_fedfuel_chart2)

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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedfuel_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_fedfuel_chart3.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fedfuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_fedfuel_1_rows):
        if not netz_fedfuel_1['fuel_code'].iloc[i] in ['Total']:
            netz_fedfuel_chart3.add_series({
                'name':       ['FED by fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, 0],
                'categories': ['FED by fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, 2,\
                    (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, netz_fedfuel_1_cols - 1],
                'values':     ['FED by fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, netz_fedfuel_1_cols - 1],
                'line':       {'color': netz_fedfuel_1['fuel_code'].map(colours_dict).loc[i], 'width': 1}
            })

        else:
            pass    
        
    ref_worksheet1.insert_chart('R' + str(chart_height + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 9), netz_fedfuel_chart3)

    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    # netz_worksheet2 = writer.sheets['FED by sector']
        
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedsector_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_fedsector_chart3.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fedsector_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_fedsector_2_rows):
        if not netz_fedsector_2['item_code_new'].iloc[i] in ['Total']:
            netz_fedsector_chart3.add_series({
                'name':       ['FED by sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, 1],
                'categories': ['FED by sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, 2, (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, netz_fedsector_2_cols - 1],
                'values':     ['FED by sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, 2, (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, netz_fedsector_2_cols - 1],
                'fill':       {'color': netz_fedsector_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
    ref_worksheet2.insert_chart('B' + str(chart_height + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 12), netz_fedsector_chart3)

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
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedsector_chart4.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_fedsector_chart4.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fedsector_chart4.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in netz_fedsector_3['item_code_new'].unique():
        i = netz_fedsector_3[netz_fedsector_3['item_code_new'] == component].index[0]
        if not netz_fedsector_3['item_code_new'].iloc[i] in ['Total']:
            netz_fedsector_chart4.add_series({
                'name':       ['FED by sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + i + 13, 1],
                'categories': ['FED by sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + 12, 2,\
                    (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + 12, netz_fedsector_3_cols - 1],
                'values':     ['FED by sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + i + 13, 2,\
                    (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + i + 13, netz_fedsector_3_cols - 1],
                'fill':       {'color': netz_fedsector_3['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })

        else:
            pass
    
    ref_worksheet2.insert_chart('J' + str(chart_height + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 12), netz_fedsector_chart4)

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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedsector_chart5.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_fedsector_chart5.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fedsector_chart5.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_fedsector_2_rows):
        if not netz_fedsector_2['item_code_new'].iloc[i] in ['Total']:
            netz_fedsector_chart5.add_series({
                'name':       ['FED by sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, 1],
                'categories': ['FED by sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, 2,\
                    (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, netz_fedsector_2_cols - 1],
                'values':     ['FED by sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, netz_fedsector_2_cols - 1],
                'line':       {'color': netz_fedsector_2['item_code_new'].map(colours_dict).loc[i], 'width': 1}
            })

        else:
            pass    
        
    ref_worksheet2.insert_chart('R' + str(chart_height + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 12), netz_fedsector_chart5)
    
    ############################# Next sheet: FED (TFC) for building sector ##################################
    
    # Access the workbook and third sheet with data from bld_df1
    # netz_worksheet3 = writer.sheets['Buildings']
    
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_fed_bld_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_fed_bld_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fed_bld_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in netz_bld_2['fuel_code'].unique():
        i = netz_bld_2[netz_bld_2['fuel_code'] == component].index[0]
        if not netz_bld_2['fuel_code'].iloc[i] in ['Total']:
            netz_fed_bld_chart1.add_series({
                'name':       ['Buildings', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + i + 7, 0],
                'categories': ['Buildings', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + 6, 2,\
                    (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + 6, netz_bld_2_cols - 1],
                'values':     ['Buildings', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + i + 7, 2,\
                    (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + i + 7, netz_bld_2_cols - 1],
                'fill':       {'color': netz_bld_2['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass

    ref_worksheet3.insert_chart('B' + str(chart_height + ref_bld_2_rows + ref_bld_3_rows + 9), netz_fed_bld_chart1)
    
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_fed_bld_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_fed_bld_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fed_bld_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for bld_sect in ['Services', 'Residential']:
        i = netz_bld_3[netz_bld_3['item_code_new'] == bld_sect].index[0]
        netz_fed_bld_chart2.add_series({
            'name':       ['Buildings', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 10 + i, 1],
            'categories': ['Buildings', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 9, 2,\
                (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 9, netz_bld_3_cols - 1],
            'values':     ['Buildings', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 10 + i, 2,\
                (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 10 + i, netz_bld_3_cols - 1],
            'fill':       {'color': netz_bld_3['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet3.insert_chart('J' + str(chart_height + ref_bld_2_rows + ref_bld_3_rows + 9), netz_fed_bld_chart2)
    
    ############################# Next sheet: FED (TFC) for industry ##################################
    
    # # Access the workbook and fourth sheet with data from bld_df1
    # netz_worksheet4 = writer.sheets['Industry']
    
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_fed_ind_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_fed_ind_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fed_ind_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_ind_1_rows):
        if not netz_ind_1['item_code_new'].iloc[i] in ['Industry']:
            netz_fed_ind_chart1.add_series({
                'name':       ['Industry', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + i + 7, 1],
                'categories': ['Industry', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + 6, 2,\
                    (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + 6, netz_ind_1_cols - 1],
                'values':     ['Industry', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + i + 7, netz_ind_1_cols - 1],
                'fill':       {'color': netz_ind_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
    ref_worksheet4.insert_chart('B' + str(chart_height + ref_ind_1_rows + ref_ind_2_rows + 9), netz_fed_ind_chart1)
    
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_fed_ind_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_fed_ind_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_fed_ind_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for fuel_agg in netz_ind_2['fuel_code'].unique():
        j = netz_ind_2[netz_ind_2['fuel_code'] == fuel_agg].index[0]
        if not netz_ind_2['fuel_code'].iloc[j] in ['Total']:
            netz_fed_ind_chart2.add_series({
                'name':       ['Industry', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + j + 10, 0],
                'categories': ['Industry', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + 9, 2,\
                    (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + 9, netz_ind_2_cols - 1],
                'values':     ['Industry', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + j + 10, 2,\
                    (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + j + 10, netz_ind_2_cols - 1],
                'fill':       {'color': netz_ind_2['fuel_code'].map(colours_dict).loc[j]},
                'border':     {'none': True}
            })

        else:
            pass
    
    ref_worksheet4.insert_chart('J' + str(chart_height + ref_ind_1_rows + ref_ind_2_rows + 9), netz_fed_ind_chart2)

    ################################# NEXT SHEET: TRANSPORT FED ################################################################

    # Access the workbook and first sheet with data from df1
    # netz_worksheet5 = writer.sheets['Transport']
        
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
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_transport_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_transport_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_transport_chart1.set_title({
        'none': True
    })
        
    for fuel_agg in netz_trn_1['fuel_code'].unique():
        j = netz_trn_1[netz_trn_1['fuel_code'] == fuel_agg].index[0]
        if not netz_trn_1['fuel_code'].iloc[j] in ['Total']:
            netz_transport_chart1.add_series({
                'name':       ['Transport', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + j + 7, 0],
                'categories': ['Transport', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + 6, 2,\
                    (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + 6, netz_trn_1_cols - 1],
                'values':     ['Transport', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + j + 7, 2,\
                    (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + j + 7, netz_trn_1_cols - 1],
                'fill':       {'color': netz_trn_1['fuel_code'].map(colours_dict).loc[j]},
                'border':     {'none': True} 
            })

        else:
            pass
    
    ref_worksheet5.insert_chart('B' + str(chart_height + ref_trn_1_rows + ref_trn_2_rows + 9), netz_transport_chart1)
            
    ############# FED transport chart 2 (transport by modality)
    
    # Create a FED transport modality column chart
    if netz_trn_2_rows > 0:
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_transport_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_transport_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_transport_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for modality in netz_trn_2['item_code_new'].unique():
            j = netz_trn_2[netz_trn_2['item_code_new'] == modality].index[0]
            if not netz_trn_2['item_code_new'].iloc[j] in ['Total']:
                netz_transport_chart2.add_series({
                    'name':       ['Transport', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + j + 10, 1],
                    'categories': ['Transport', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + 9, 2,\
                        (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + 9, netz_trn_2_cols - 1],
                    'values':     ['Transport', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + j + 10, 2,\
                        (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + j + 10, netz_trn_2_cols - 1],
                    'fill':       {'color': netz_trn_2['item_code_new'].map(colours_dict).loc[j]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        ref_worksheet5.insert_chart('J' + str(chart_height + ref_trn_1_rows + ref_trn_2_rows + 9), netz_transport_chart2)

    else:
        pass

    ################################# NEXT SHEET: AGRICULTURE FED ################################################################

    # Access the workbook and first sheet with data from df1
    # netz_worksheet6 = writer.sheets['Agriculture']
        
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet6.set_column(2, netz_ag_1_cols + 1, None, space_format)
    # netz_worksheet6.set_row(chart_height, None, header_format)
    # netz_worksheet6.set_row(chart_height + netz_ag_1_rows + 3, None, header_format)
    # netz_worksheet6.write(0, 0, economy + ' FED agriculture', cell_format1)

    # Create a Agriculture line chart 
    if netz_ag_1_rows > 0:
        netz_ag_chart1 = workbook.add_chart({'type': 'line'})
        netz_ag_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_ag_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_ag_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_ag_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_ag_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_ag_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_ag_1_rows):
            if not netz_ag_1['fuel_code'].iloc[i] in ['Total']:
                netz_ag_chart1.add_series({
                    'name':       ['Agriculture', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, 0],
                    'categories': ['Agriculture', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, 2,\
                        (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, netz_ag_1_cols - 1],
                    'values':     ['Agriculture', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, netz_ag_1_cols - 1],
                    'line':       {'color': netz_ag_1['fuel_code'].map(colours_dict).loc[i], 'width': 1}
                })

            else:
                pass    
            
        ref_worksheet6.insert_chart('B' + str(chart_height + ref_ag_1_rows + ref_ag_2_rows + 9), netz_ag_chart1)

    else:
        pass

    # Create a Agriculture area chart
    if netz_ag_1_rows > 0:
        netz_ag_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_ag_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_ag_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_ag_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_ag_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_ag_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_ag_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_ag_1_rows):
            if not netz_ag_1['fuel_code'].iloc[i] in ['Total']:
                netz_ag_chart2.add_series({
                    'name':       ['Agriculture', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, 0],
                    'categories': ['Agriculture', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, 2,\
                        (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, netz_ag_1_cols - 1],
                    'values':     ['Agriculture', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, netz_ag_1_cols - 1],
                    'fill':       {'color': netz_ag_1['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet6.insert_chart('J' + str(chart_height + ref_ag_1_rows + ref_ag_2_rows + 9), netz_ag_chart2)

    else:
        pass

    # Create a Agriculture stacked column
    if netz_ag_2_rows > 0:
        netz_ag_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
        netz_ag_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_ag_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        netz_ag_chart3.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'interval_unit': 1,
            'line': {'color': '#bebebe'}
        })
            
        netz_ag_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_ag_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_ag_chart3.set_title({
            'none': True
        })

        # Configure the series of the chart from the dataframe data.    
        for i in range(netz_ag_2_rows):
            if not netz_ag_2['fuel_code'].iloc[i] in ['Total']:
                netz_ag_chart3.add_series({
                    'name':       ['Agriculture', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + i + 10, 0],
                    'categories': ['Agriculture', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + 9, 2,\
                        (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + 9, netz_ag_2_cols - 1],
                    'values':     ['Agriculture', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + i + 10,\
                        2, (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + i + 10, netz_ag_2_cols - 1],
                    'fill':       {'color': netz_ag_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        ref_worksheet6.insert_chart('R' + str(chart_height + ref_ag_1_rows + ref_ag_2_rows + 9), netz_ag_chart3)

    else:
        pass

    ############################################################################################################################

    # TPES charts

    ################################################################### CHARTS ###################################################################
    # REFERENCE
    # Access the sheet created using writer above
    ref_worksheet11 = writer.sheets['Supply']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet11.set_column(2, ref_tpes_1_cols + 1, None, space_format)
    ref_worksheet11.set_row(chart_height, None, header_format)
    ref_worksheet11.set_row(chart_height + ref_tpes_1_rows + 3, None, header_format)
    ref_worksheet11.set_row((2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, None, header_format)
    ref_worksheet11.set_row((2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + ref_tpes_1_rows + 9, None, header_format)
    ref_worksheet11.write(0, 0, economy + ' TPES fuel Reference', cell_format1)
    ref_worksheet11.write(chart_height + ref_tpes_1_rows + ref_tpes_2_rows + 6, 0, economy + ' TPES fuel Carbon Neutrality', cell_format1)
    ref_worksheet11.write(1, 0, 'Units: Petajoules', cell_format2)

    ######################################################
    # Create a TPES chart
    if ref_tpes_1_rows > 0:
        ref_tpes_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_tpes_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_tpes_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_tpes_1_rows):
            if not ref_tpes_1['fuel_code'].iloc[i] in ['Total']:
                ref_tpes_chart2.add_series({
                    'name':       ['Supply', chart_height + i + 1, 0],
                    'categories': ['Supply', chart_height, 2, chart_height, ref_tpes_1_cols - 1],
                    'values':     ['Supply', chart_height + i + 1, 2, chart_height + i + 1, ref_tpes_1_cols - 1],
                    'fill':       {'color': ref_tpes_1['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })    
            
            else:
                pass

        ref_worksheet11.insert_chart('B3', ref_tpes_chart2)

    else:
        pass

    ######## same chart as above but line

    # TPES line chart
    if ref_tpes_1_rows > 0:
        ref_tpes_chart4 = workbook.add_chart({'type': 'line'})
        ref_tpes_chart4.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_chart4.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_chart4.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_chart4.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_tpes_chart4.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_chart4.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_tpes_1_rows):
            if not ref_tpes_1['fuel_code'].iloc[i] in ['Total']:
                ref_tpes_chart4.add_series({
                    'name':       ['Supply', chart_height + i + 1, 0],
                    'categories': ['Supply', chart_height, 2, chart_height, ref_tpes_1_cols - 1],
                    'values':     ['Supply', chart_height + i + 1, 2, chart_height + i + 1, ref_tpes_1_cols - 1],
                    'line':       {'color': ref_tpes_1['fuel_code'].map(colours_dict).loc[i], 
                                'width': 1}
                })

            else:
                ref_tpes_chart4.add_series({
                    'name':       ['Supply', chart_height + i + 1, 0],
                    'categories': ['Supply', chart_height, 2, chart_height, ref_tpes_1_cols - 1],
                    'values':     ['Supply', chart_height + i + 1, 2, chart_height + i + 1, ref_tpes_1_cols - 1],
                    'line':       {'color': ref_tpes_1['fuel_code'].map(colours_dict).loc[i], 
                                'width': 1.5}
                })    
            
        ref_worksheet11.insert_chart('R3', ref_tpes_chart4)

    else:
        pass

    ###################### Create another TPES chart showing proportional share #################################

    # Create a TPES chart
    if ref_tpes_2_rows > 0:
        ref_tpes_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_chart3.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'interval_unit': 1,
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_chart3.set_title({
            'none': True
        })

        # Configure the series of the chart from the dataframe data.    
        for component in ref_tpes_2['fuel_code'].unique():
            i = ref_tpes_2[ref_tpes_2['fuel_code'] == component].index[0]
            if not ref_tpes_2['fuel_code'].iloc[i] in ['Total']:
                ref_tpes_chart3.add_series({
                    'name':       ['Supply', chart_height + ref_tpes_1_rows + i + 4, 0],
                    'categories': ['Supply', chart_height + ref_tpes_1_rows + 3, 2, chart_height + ref_tpes_1_rows + 3, ref_tpes_2_cols - 1],
                    'values':     ['Supply', chart_height + ref_tpes_1_rows + i + 4, 2, chart_height + ref_tpes_1_rows + i + 4, ref_tpes_2_cols - 1],
                    'fill':       {'color': ref_tpes_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        ref_worksheet11.insert_chart('J3', ref_tpes_chart3)

    else:
        pass

    ########################################### PRODUCTION CHARTS #############################################
    
    # access the sheet for production created above
    ref_worksheet12 = writer.sheets['Production']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet12.set_column(2, ref_prod_1_cols + 1, None, space_format)
    ref_worksheet12.set_row(chart_height, None, header_format)
    ref_worksheet12.set_row(chart_height + ref_prod_1_rows + 3, None, header_format)
    ref_worksheet12.set_row((2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, None, header_format)
    ref_worksheet12.set_row((2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + 9, None, header_format)
    ref_worksheet12.write(0, 0, economy + ' production by fuel Reference', cell_format1)
    ref_worksheet12.write(chart_height + ref_prod_1_rows + ref_prod_2_rows + 6, 0, economy + ' production by fuel Carbon Neutrality', cell_format1)
    ref_worksheet12.write(1, 0, 'Units: Petajoules', cell_format2)    

    ###################### Create another PRODUCTION chart with only 6 categories #################################

    # Create a PROD chart
    if ref_prod_1_rows > 0:
        ref_prod_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_prod_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_prod_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        ref_prod_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_prod_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_prod_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_prod_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_prod_1_rows):
            if not ref_prod_1['fuel_code'].iloc[i] in ['Total']:
                ref_prod_chart2.add_series({
                    'name':       ['Production', chart_height + i + 1, 0],
                    'categories': ['Production', chart_height, 2, chart_height, ref_prod_1_cols - 1],
                    'values':     ['Production', chart_height + i + 1, 2, chart_height + i + 1, ref_prod_1_cols - 1],
                    'fill':       {'color': ref_prod_1['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet12.insert_chart('B3', ref_prod_chart2)

    else:
        pass

    ############ Same as above but with a line ###########

    # Create a PROD chart
    if ref_prod_1_rows > 0:
        ref_prod_chart2 = workbook.add_chart({'type': 'line'})
        ref_prod_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_prod_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        ref_prod_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_prod_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_prod_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_prod_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_prod_1_rows):
            if not ref_prod_1['fuel_code'].iloc[i] in ['Total']:
                ref_prod_chart2.add_series({
                    'name':       ['Production', chart_height + i + 1, 0],
                    'categories': ['Production', chart_height, 2, chart_height, ref_prod_1_cols - 1],
                    'values':     ['Production', chart_height + i + 1, 2, chart_height + i + 1, ref_prod_1_cols - 1],
                    'line':       {'color': ref_prod_1['fuel_code'].map(colours_dict).loc[i],
                                'width': 1} 
                })

            else:
                pass    
            
        ref_worksheet12.insert_chart('R3', ref_prod_chart2)

    else:
        pass

    ###################### Create another PRODUCTION chart showing proportional share #################################

    # Create a production chart
    if ref_prod_2_rows > 0:
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'interval_unit': 1,
            'line': {'color': '#bebebe'}
        })
            
        ref_prod_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_prod_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_prod_chart3.set_title({
            'none': True
        })

        # Configure the series of the chart from the dataframe data.    
        for component in ref_prod_2['fuel_code'].unique():
            i = ref_prod_2[ref_prod_2['fuel_code'] == component].index[0]
            if not ref_prod_2['fuel_code'].iloc[i] in ['Total']:
                ref_prod_chart3.add_series({
                    'name':       ['Production', chart_height + ref_prod_1_rows + i + 4, 0],
                    'categories': ['Production', chart_height + ref_prod_1_rows + 3, 2, chart_height + ref_prod_1_rows + 3, ref_prod_2_cols - 1],
                    'values':     ['Production', chart_height + ref_prod_1_rows + i + 4, 2, chart_height + ref_prod_1_rows + i + 4, ref_prod_2_cols - 1],
                    'fill':       {'color': ref_prod_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        ref_worksheet12.insert_chart('J3', ref_prod_chart3)

    else:
        pass
    
    ###################################### TPES components I ###########################################
    
    # access the sheet for production created above
    ref_worksheet13 = writer.sheets['Supply REF']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet13.set_column(2, ref_imports_1_cols + 1, None, space_format)
    ref_worksheet13.set_row(chart_height, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + 3, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + 6, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + 9, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + 12, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_1_rows + 15, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_1_rows + ref_electrade_1_rows + 18, None, header_format)
    ref_worksheet13.write(0, 0, economy + ' TPES components Reference', cell_format1)
    ref_worksheet13.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a TPES components chart
    if ref_tpes_comp_1_rows > 0:
        ref_tpes_comp_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_comp_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_comp_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_comp_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_comp_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_comp_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_comp_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in ref_tpes_comp_1['item_code_new'].unique():
            i = ref_tpes_comp_1[ref_tpes_comp_1['item_code_new'] == component].index[0]
            ref_tpes_comp_chart1.add_series({
                'name':       ['Supply REF', chart_height + i + 1, 1],
                'categories': ['Supply REF', chart_height, 2, chart_height, ref_tpes_comp_1_cols - 1],
                'values':     ['Supply REF', chart_height + i + 1, 2, chart_height + i + 1, ref_tpes_comp_1_cols - 1],
                'fill':       {'color': ref_tpes_comp_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet13.insert_chart('B3', ref_tpes_comp_chart1)

    else: 
        pass

    # IMPORTS: Create a line chart subset by fuel
    
    if ref_imports_1_rows > 0:
        ref_imports_line = workbook.add_chart({'type': 'line'})
        ref_imports_line.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_imports_line.set_chartarea({
            'border': {'none': True}
        })
        
        ref_imports_line.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_imports_line.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Imports (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_imports_line.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_imports_line.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for fuel in ref_imports_1['fuel_code'].unique():
            i = ref_imports_1[ref_imports_1['fuel_code'] == fuel].index[0]
            if not ref_imports_1['fuel_code'].iloc[i] in ['Total']:
                ref_imports_line.add_series({
                    'name':       ['Supply REF', chart_height + ref_tpes_comp_1_rows + i + 4, 0],
                    'categories': ['Supply REF', chart_height + ref_tpes_comp_1_rows + 3, 2, chart_height + ref_tpes_comp_1_rows + 3, ref_imports_1_cols - 1],
                    'values':     ['Supply REF', chart_height + ref_tpes_comp_1_rows + i + 4, 2, chart_height + ref_tpes_comp_1_rows + i + 4, ref_imports_1_cols - 1],
                    'line':       {'color': ref_imports_1['fuel_code'].map(colours_dict).loc[i], 
                                'width': 1},
                })

            else:
                pass    
            
        ref_worksheet13.insert_chart('J3', ref_imports_line)

    else:
        pass

    # Create a imports by fuel column
    if ref_imports_2_rows > 0:
        ref_imports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_imports_column.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_imports_column.set_chartarea({
            'border': {'none': True}
        })
        
        ref_imports_column.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_imports_column.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Imports by fuel (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_imports_column.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_imports_column.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_imports_2_rows):
            if not ref_imports_2['fuel_code'].iloc[i] in ['Total']:
                ref_imports_column.add_series({
                    'name':       ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + i + 7, 0],
                    'categories': ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + 6, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + 6, ref_imports_2_cols - 1],
                    'values':     ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + i + 7, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + i + 7, ref_imports_2_cols - 1],
                    'fill':       {'color': ref_imports_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        ref_worksheet13.insert_chart('R3', ref_imports_column)

    else:
        pass

    # EXPORTS: Create a line chart subset by fuel
    
    if ref_exports_1_rows > 0:
        ref_exports_line = workbook.add_chart({'type': 'line'})
        ref_exports_line.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_exports_line.set_chartarea({
            'border': {'none': True}
        })
        
        ref_exports_line.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_exports_line.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Exports (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_exports_line.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_exports_line.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for fuel in ref_exports_1['fuel_code'].unique():
            i = ref_exports_1[ref_exports_1['fuel_code'] == fuel].index[0]
            if not ref_exports_1['fuel_code'].iloc[i] in ['Total']:
                ref_exports_line.add_series({
                    'name':       ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + i + 10, 0],
                    'categories': ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + 9, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + 9, ref_imports_1_cols - 1],
                    'values':     ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + i + 10, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + i + 10, ref_imports_1_cols - 1],
                    'line':       {'color': ref_exports_1['fuel_code'].map(colours_dict).loc[i], 
                                'width': 1},
                })

            else:
                pass    

        # 40    
        ref_worksheet13.insert_chart('Z3', ref_exports_line)

    else:
        pass

    # Create a imports by fuel column
    if ref_exports_2_rows > 0:
        ref_exports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_exports_column.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_exports_column.set_chartarea({
            'border': {'none': True}
        })
        
        ref_exports_column.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_exports_column.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Exports by fuel (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_exports_column.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_exports_column.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_exports_2_rows):
            if not ref_exports_2['fuel_code'].iloc[i] in ['Total']:
                ref_exports_column.add_series({
                    'name':       ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + i + 13, 0],
                    'categories': ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + 12, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + 12, ref_exports_2_cols - 1],
                    'values':     ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + i + 13, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + i + 13, ref_exports_2_cols - 1],
                    'fill':       {'color': ref_exports_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        ref_worksheet13.insert_chart('AH3', ref_exports_column)

    else:
        pass

    # Create an electricity trade column
    if ref_electrade_1_rows > 0:
        ref_electrade_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_electrade_column.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_electrade_column.set_chartarea({
            'border': {'none': True}
        })
        
        ref_electrade_column.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_electrade_column.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Electricity trade (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_electrade_column.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_electrade_column.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_electrade_1_rows):
            ref_electrade_column.add_series({
                'name':       ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + i + 16, 1],
                'categories': ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + 15, 2,\
                    chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + 15, ref_electrade_1_cols - 1],
                'values':     ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + i + 16, 2,\
                    chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_1_rows + i + 16, ref_electrade_1_cols - 1],
                'fill':       {'color': ref_electrade_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet13.insert_chart('AP3', ref_electrade_column)

    else:
        pass

    # Create a net trade column
    if ref_nettrade_1_rows > 0:
        ref_nettrade_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_nettrade_column.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_nettrade_column.set_chartarea({
            'border': {'none': True}
        })
        
        ref_nettrade_column.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_nettrade_column.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Net trade (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_nettrade_column.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_nettrade_column.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_nettrade_1_rows):
            if not ref_nettrade_1['fuel_code'].iloc[i] in ['Trade balance']:
                ref_nettrade_column.add_series({
                    'name':       ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + ref_electrade_1_rows + i + 19, 0],
                    'categories': ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + ref_electrade_1_rows + 18, 2,\
                        chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + ref_electrade_1_rows + 18, ref_nettrade_1_cols - 1],
                    'values':     ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + ref_electrade_1_rows + i + 19, 2,\
                        chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_1_rows + ref_electrade_1_rows + i + 19, ref_nettrade_1_cols - 1],
                    'fill':       {'color': ref_nettrade_1['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass

        # Marker series
        ref_nettrade_line = workbook.add_chart({'type': 'line'})
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_nettrade_1_rows):
            if ref_nettrade_1['fuel_code'].iloc[i] in ['Trade balance']:
                ref_nettrade_line.add_series({
                    'name':       ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + ref_electrade_1_rows + i + 19, 0],
                    'categories': ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + ref_electrade_1_rows + 18, 2,\
                        chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + ref_electrade_1_rows + 18, ref_nettrade_1_cols - 1],
                    'values':     ['Supply REF', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_2_rows + ref_electrade_1_rows + i + 19, 2,\
                        chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + ref_exports_1_rows + ref_electrade_1_rows + i + 19, ref_nettrade_1_cols - 1],
                    'marker':     {'fill': {'color': ref_nettrade_1['fuel_code'].map(colours_dict).loc[i]},
                                   'type': 'diamond',
                                   'border': {'none': True},
                                   'size': 8},
                    'line':       {'none': True}
                })

            else:
                pass

        ref_nettrade_column.combine(ref_nettrade_line)
        
        ref_worksheet13.insert_chart('AX3', ref_nettrade_column)

    else:
        pass

    ###################################### TPES components II ###########################################
    
    # access the sheet for production created above
    ref_worksheet14 = writer.sheets['Bunkers']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet14.set_column(2, ref_bunkers_1_cols + 1, None, space_format)
    ref_worksheet14.set_row(chart_height, None, header_format)
    ref_worksheet14.set_row(chart_height + ref_bunkers_1_rows + 3, None, header_format)
    ref_worksheet14.set_row((2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + 6, None, header_format)
    ref_worksheet14.set_row((2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + 9, None, header_format)
    ref_worksheet14.write(0, 0, economy + ' TPES bunkers Reference', cell_format1)
    ref_worksheet14.write(chart_height + ref_bunkers_1_rows + ref_bunkers_2_rows + 6, 0, economy + ' TPES bunkers Carbon Neutrality', cell_format1)
    ref_worksheet14.write(1, 0, 'Units: Petajoules', cell_format2)

    # MARINE BUNKER: Create a line chart subset by fuel
    if ref_bunkers_1_rows > 0:
        ref_marine_line = workbook.add_chart({'type': 'line'})
        ref_marine_line.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_marine_line.set_chartarea({
            'border': {'none': True}
        })
        
        ref_marine_line.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_marine_line.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Marine bunkers (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_marine_line.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_marine_line.set_title({
            'none': True
        }) 

        # Configure the series of the chart from the dataframe data.
        for i in range(ref_bunkers_1_rows):
            ref_marine_line.add_series({
                'name':       ['Bunkers', chart_height + i + 1, 0],
                'categories': ['Bunkers', chart_height, 2, chart_height, ref_bunkers_1_cols - 1],
                'values':     ['Bunkers', chart_height + i + 1, 2, chart_height + i + 1, ref_bunkers_1_cols - 1],
                'line':       {'color': ref_bunkers_1['fuel_code'].map(colours_dict).loc[i], 
                            'width': 1},
            })    
            
        ref_worksheet14.insert_chart('B3', ref_marine_line)

    else:
        pass

    # AVIATION BUNKER: Create a line chart subset by fuel
    if ref_bunkers_2_rows > 0:
        ref_aviation_line = workbook.add_chart({'type': 'line'})
        ref_aviation_line.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_aviation_line.set_chartarea({
            'border': {'none': True}
        })
        
        ref_aviation_line.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_aviation_line.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Aviation bunkers (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_aviation_line.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_aviation_line.set_title({
            'none': True
        }) 

        # Configure the series of the chart from the dataframe data.
        for i in range(ref_bunkers_2_rows):
            ref_aviation_line.add_series({
                'name':       ['Bunkers', chart_height + ref_bunkers_1_rows + i + 4, 0],
                'categories': ['Bunkers', chart_height + ref_bunkers_1_rows + 3, 2, chart_height + ref_bunkers_1_rows + 3, ref_bunkers_2_cols - 1],
                'values':     ['Bunkers', chart_height + ref_bunkers_1_rows + i + 4, 2, chart_height + ref_bunkers_1_rows + i + 4, ref_bunkers_2_cols - 1],
                'line':       {'color': ref_bunkers_2['fuel_code'].map(colours_dict).loc[i], 
                            'width': 1},
            })    
            
        ref_worksheet14.insert_chart('J3', ref_aviation_line)
    
    else:
        pass

    ###############################################################################################################
    # CARBON NEUTRALITY
    # Access the sheet created using writer above
    # netz_worksheet11 = writer.sheets['TPES']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet11.set_column(2, netz_tpes_1_cols + 1, None, space_format)
    # netz_worksheet11.set_row(chart_height, None, header_format)
    # netz_worksheet11.set_row(chart_height + netz_tpes_1_rows + 3, None, header_format)
    # netz_worksheet11.write(0, 0, economy + ' TPES fuel net-zero', cell_format1)

    ######################################################
    # Create a TPES chart
    if netz_tpes_1_rows > 0:
        netz_tpes_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_tpes_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_tpes_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_tpes_1_rows):
            if not netz_tpes_1['fuel_code'].iloc[i] in ['Total']:
                netz_tpes_chart2.add_series({
                    'name':       ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 0],
                    'categories': ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, 2,\
                        (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, netz_tpes_1_cols - 1],
                    'values':     ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, netz_tpes_1_cols - 1],
                    'fill':       {'color': netz_tpes_1['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet11.insert_chart('B' + str(chart_height + ref_tpes_1_rows + ref_tpes_2_rows + 9), netz_tpes_chart2)

    else:
        pass

    ######## same chart as above but line

    # TPES line chart
    if netz_tpes_1_rows > 0:
        netz_tpes_chart4 = workbook.add_chart({'type': 'line'})
        netz_tpes_chart4.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_chart4.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_chart4.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_chart4.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_tpes_chart4.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_chart4.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_tpes_1_rows):
            if not netz_tpes_1['fuel_code'].iloc[i] in ['Total']:
                netz_tpes_chart4.add_series({
                    'name':       ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 0],
                    'categories': ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, 2,\
                        (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, netz_tpes_1_cols - 1],
                    'values':     ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, netz_tpes_1_cols - 1],
                    'line':       {'color': netz_tpes_1['fuel_code'].map(colours_dict).loc[i], 
                                'width': 1}
                })

            else:
                netz_tpes_chart4.add_series({
                    'name':       ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 0],
                    'categories': ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, 2,\
                        (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, netz_tpes_1_cols - 1],
                    'values':     ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, netz_tpes_1_cols - 1],
                    'line':       {'color': netz_tpes_1['fuel_code'].map(colours_dict).loc[i], 
                                'width': 1.5}
                })    
            
        ref_worksheet11.insert_chart('R' + str(chart_height + ref_tpes_1_rows + ref_tpes_2_rows + 9), netz_tpes_chart4)

    else:
        pass

    ###################### Create another TPES chart showing proportional share #################################

    # Create a TPES chart
    if netz_tpes_2_rows > 0:
        netz_tpes_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_chart3.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'interval_unit': 1,
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_chart3.set_title({
            'none': True
        })

        # Configure the series of the chart from the dataframe data.    
        for component in netz_tpes_2['fuel_code'].unique():
            i = netz_tpes_2[netz_tpes_2['fuel_code'] == component].index[0]
            if not netz_tpes_2['fuel_code'].iloc[i] in ['Total']:
                netz_tpes_chart3.add_series({
                    'name':       ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + i + 10, 0],
                    'categories': ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + 9, 2,\
                        (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + 9, netz_tpes_2_cols - 1],
                    'values':     ['Supply', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + i + 10, 2,\
                        (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + i + 10, netz_tpes_2_cols - 1],
                    'fill':       {'color': netz_tpes_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        ref_worksheet11.insert_chart('J' + str(chart_height + ref_tpes_1_rows + ref_tpes_2_rows + 9), netz_tpes_chart3)

    else:
        pass

    ########################################### PRODUCTION CHARTS #############################################
    
    # access the sheet for production created above
    # netz_worksheet12 = writer.sheets['Production']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet12.set_column(2, netz_prod_1_cols + 1, None, space_format)
    # netz_worksheet12.set_row(chart_height, None, header_format)
    # netz_worksheet12.set_row(chart_height + netz_prod_1_rows + 3, None, header_format)
    # netz_worksheet12.write(0, 0, economy + ' prod fuel net-zero', cell_format1)

    ###################### Create another PRODUCTION chart with only 6 categories #################################

    # Create a PROD chart
    if netz_prod_1_rows > 0:
        netz_prod_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_prod_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_prod_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_prod_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_prod_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_prod_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_prod_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_prod_1_rows):
            if not netz_prod_1['fuel_code'].iloc[i] in ['Total']:
                netz_prod_chart2.add_series({
                    'name':       ['Production', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, 0],
                    'categories': ['Production', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, 2,\
                        (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, netz_prod_1_cols - 1],
                    'values':     ['Production', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, netz_prod_1_cols - 1],
                    'fill':       {'color': netz_prod_1['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet12.insert_chart('B' + str(chart_height + ref_prod_1_rows + ref_prod_2_rows + 9), netz_prod_chart2)

    else:
        pass

    ############ Same as above but with a line ###########

    # Create a PROD chart
    if netz_prod_1_rows > 0:
        netz_prod_chart2 = workbook.add_chart({'type': 'line'})
        netz_prod_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_prod_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_prod_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_prod_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_prod_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_prod_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_prod_1_rows):
            if not netz_prod_1['fuel_code'].iloc[i] in ['Total']:
                netz_prod_chart2.add_series({
                    'name':       ['Production', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, 0],
                    'categories': ['Production', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, 2,\
                        (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, netz_prod_1_cols - 1],
                    'values':     ['Production', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, netz_prod_1_cols - 1],
                    'line':       {'color': netz_prod_1['fuel_code'].map(colours_dict).loc[i],
                                'width': 1} 
                })

            else:
                pass    
            
        ref_worksheet12.insert_chart('R' + str(chart_height + ref_prod_1_rows + ref_prod_2_rows + 9), netz_prod_chart2)

    else:
        pass

    ###################### Create another PRODUCTION chart showing proportional share #################################

    # Create a production chart
    if netz_prod_2_rows > 0:
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'interval_unit': 1,
            'line': {'color': '#bebebe'}
        })
            
        netz_prod_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_prod_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_prod_chart3.set_title({
            'none': True
        })

        # Configure the series of the chart from the dataframe data.    
        for component in netz_prod_2['fuel_code'].unique():
            i = netz_prod_2[netz_prod_2['fuel_code'] == component].index[0]
            if not netz_prod_2['fuel_code'].iloc[i] in ['Total']:
                netz_prod_chart3.add_series({
                    'name':       ['Production', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + i + 10, 0],
                    'categories': ['Production', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + 9, 2,\
                        (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + 9, netz_prod_2_cols - 1],
                    'values':     ['Production', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + i + 10, 2,\
                        (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + i + 10, netz_prod_2_cols - 1],
                    'fill':       {'color': netz_prod_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        ref_worksheet12.insert_chart('J' + str(chart_height + ref_prod_1_rows + ref_prod_2_rows + 9), netz_prod_chart3)

    else:
        pass
    
    ###################################### TPES components I ###########################################
    
    # access the sheet for production created above
    netz_worksheet13 = writer.sheets['Supply CN']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet13.set_column(2, netz_imports_1_cols + 1, None, space_format)
    netz_worksheet13.set_row(chart_height, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + 3, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + 6, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + 9, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + 12, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_1_rows + 15, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_1_rows + netz_electrade_1_rows + 18, None, header_format)
    netz_worksheet13.write(0, 0, economy + ' TPES components Carbon Neutrality', cell_format1)
    netz_worksheet13.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a TPES components chart
    if netz_tpes_comp_1_rows > 0:
        netz_tpes_comp_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_comp_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_comp_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_comp_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_comp_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_comp_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_comp_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in netz_tpes_comp_1['item_code_new'].unique():
            i = netz_tpes_comp_1[netz_tpes_comp_1['item_code_new'] == component].index[0]
            netz_tpes_comp_chart1.add_series({
                'name':       ['Supply CN', chart_height + i + 1, 1],
                'categories': ['Supply CN', chart_height, 2, chart_height, netz_tpes_comp_1_cols - 1],
                'values':     ['Supply CN', chart_height + i + 1, 2, chart_height + i + 1, netz_tpes_comp_1_cols - 1],
                'fill':       {'color': netz_tpes_comp_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        netz_worksheet13.insert_chart('B3', netz_tpes_comp_chart1)

    else:
        pass

    # IMPORTS: Create a line chart subset by fuel
    if netz_imports_1_rows > 0:
        netz_imports_line = workbook.add_chart({'type': 'line'})
        netz_imports_line.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_imports_line.set_chartarea({
            'border': {'none': True}
        })
        
        netz_imports_line.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_imports_line.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Imports (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_imports_line.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_imports_line.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for fuel in netz_imports_1['fuel_code'].unique():
            i = netz_imports_1[netz_imports_1['fuel_code'] == fuel].index[0]
            if not netz_imports_1['fuel_code'].iloc[i] in ['Total']:
                netz_imports_line.add_series({
                    'name':       ['Supply CN', chart_height + netz_tpes_comp_1_rows + i + 4, 0],
                    'categories': ['Supply CN', chart_height + netz_tpes_comp_1_rows + 3, 2, chart_height + netz_tpes_comp_1_rows + 3, netz_imports_1_cols - 1],
                    'values':     ['Supply CN', chart_height + netz_tpes_comp_1_rows + i + 4, 2, chart_height + netz_tpes_comp_1_rows + i + 4, netz_imports_1_cols - 1],
                    'line':       {'color': netz_imports_1['fuel_code'].map(colours_dict).loc[i], 
                                'width': 1},
                })

            else:
                pass    
            
        netz_worksheet13.insert_chart('J3', netz_imports_line)

    else:
        pass

    # Create a imports by fuel column
    if netz_imports_2_rows > 0:
        netz_imports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_imports_column.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_imports_column.set_chartarea({
            'border': {'none': True}
        })
        
        netz_imports_column.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_imports_column.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Imports by fuel (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_imports_column.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_imports_column.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(netz_imports_2_rows):
            if not netz_imports_2['fuel_code'].iloc[i] in ['Total']:
                netz_imports_column.add_series({
                    'name':       ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + i + 7, 0],
                    'categories': ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + 6, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + 6, netz_imports_2_cols - 1],
                    'values':     ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + i + 7, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + i + 7, netz_imports_2_cols - 1],
                    'fill':       {'color': netz_imports_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        netz_worksheet13.insert_chart('R3', netz_imports_column)

    else:
        pass

    # EXPORTS: Create a line chart subset by fuel
    if netz_exports_1_rows > 0:
        netz_exports_line = workbook.add_chart({'type': 'line'})
        netz_exports_line.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_exports_line.set_chartarea({
            'border': {'none': True}
        })
        
        netz_exports_line.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_exports_line.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Exports (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_exports_line.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_exports_line.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for fuel in netz_exports_1['fuel_code'].unique():
            i = netz_exports_1[netz_exports_1['fuel_code'] == fuel].index[0]
            if not netz_exports_1['fuel_code'].iloc[i] in ['Total']:
                netz_exports_line.add_series({
                    'name':       ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + i + 10, 0],
                    'categories': ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + 9, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + 9, netz_imports_1_cols - 1],
                    'values':     ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + i + 10, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + i + 10, netz_imports_1_cols - 1],
                    'line':       {'color': netz_exports_1['fuel_code'].map(colours_dict).loc[i], 
                                'width': 1},
                })

            else:
                pass    
            
        netz_worksheet13.insert_chart('Z3', netz_exports_line)

    else:
        pass

    # Create a imports by fuel column
    if netz_exports_2_rows > 0:
        netz_exports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_exports_column.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_exports_column.set_chartarea({
            'border': {'none': True}
        })
        
        netz_exports_column.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_exports_column.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Exports by fuel (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_exports_column.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_exports_column.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(netz_exports_2_rows):
            if not netz_exports_2['fuel_code'].iloc[i] in ['Total']:
                netz_exports_column.add_series({
                    'name':       ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + i + 13, 0],
                    'categories': ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + 12, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + 12, netz_exports_2_cols - 1],
                    'values':     ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + i + 13, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + i + 13, netz_exports_2_cols - 1],
                    'fill':       {'color': netz_exports_2['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass
        
        netz_worksheet13.insert_chart('AH3', netz_exports_column)

    else:
        pass

    # Create an electricity trade column
    if netz_electrade_1_rows > 0:
        netz_electrade_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_electrade_column.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_electrade_column.set_chartarea({
            'border': {'none': True}
        })
        
        netz_electrade_column.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_electrade_column.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Electricity trade (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_electrade_column.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_electrade_column.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(netz_electrade_1_rows):
            netz_electrade_column.add_series({
                'name':       ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + i + 16, 1],
                'categories': ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + 15, 2,\
                    chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + 15, netz_electrade_1_cols - 1],
                'values':     ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + i + 16, 2,\
                    chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_1_rows + i + 16, netz_electrade_1_cols - 1],
                'fill':       {'color': netz_electrade_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        netz_worksheet13.insert_chart('AP3', netz_electrade_column)

    else:
        pass

    # Create a net trade column
    if netz_nettrade_1_rows > 0:
        netz_nettrade_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_nettrade_column.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_nettrade_column.set_chartarea({
            'border': {'none': True}
        })
        
        netz_nettrade_column.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_nettrade_column.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Net trade (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_nettrade_column.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_nettrade_column.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(netz_nettrade_1_rows):
            if not netz_nettrade_1['fuel_code'].iloc[i] in ['Trade balance']:
                netz_nettrade_column.add_series({
                    'name':       ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + netz_electrade_1_rows + i + 19, 0],
                    'categories': ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + netz_electrade_1_rows + 18, 2,\
                        chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + netz_electrade_1_rows + 18, netz_nettrade_1_cols - 1],
                    'values':     ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + netz_electrade_1_rows + i + 19, 2,\
                        chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_1_rows + netz_electrade_1_rows + i + 19, netz_nettrade_1_cols - 1],
                    'fill':       {'color': netz_nettrade_1['fuel_code'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass

        # Marker series
        netz_nettrade_line = workbook.add_chart({'type': 'line'})
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(netz_nettrade_1_rows):
            if netz_nettrade_1['fuel_code'].iloc[i] in ['Trade balance']:
                netz_nettrade_line.add_series({
                    'name':       ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + netz_electrade_1_rows + i + 19, 0],
                    'categories': ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + netz_electrade_1_rows + 18, 2,\
                        chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + netz_electrade_1_rows + 18, netz_nettrade_1_cols - 1],
                    'values':     ['Supply CN', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_2_rows + netz_electrade_1_rows + i + 19, 2,\
                        chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + netz_exports_1_rows + netz_electrade_1_rows + i + 19, netz_nettrade_1_cols - 1],
                    'marker':     {'fill': {'color': netz_nettrade_1['fuel_code'].map(colours_dict).loc[i]},
                                   'type': 'diamond',
                                   'border': {'none': True},
                                   'size': 8},
                    'line':       {'none': True}
                })

            else:
                pass

        netz_nettrade_column.combine(netz_nettrade_line)
        
        netz_worksheet13.insert_chart('AX3', netz_nettrade_column)

    else:
        pass

    ###################################### TPES components II ###########################################
    
    # access the sheet for production created above
    # netz_worksheet14 = writer.sheets[economy + 'bunkers']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet14.set_column(2, netz_bunkers_1_cols + 1, None, space_format)
    # netz_worksheet14.set_row(chart_height, None, header_format)
    # netz_worksheet14.set_row(chart_height + netz_bunkers_1_rows + 3, None, header_format)
    # netz_worksheet14.write(0, 0, economy + ' TPES components II net-zero', cell_format1)
    
    # MARINE BUNKER: Create a line chart subset by fuel
    if netz_bunkers_1_rows > 0:
        netz_marine_line = workbook.add_chart({'type': 'line'})
        netz_marine_line.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_marine_line.set_chartarea({
            'border': {'none': True}
        })
        
        netz_marine_line.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_marine_line.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Marine bunkers (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_marine_line.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_marine_line.set_title({
            'none': True
        }) 

        # Configure the series of the chart from the dataframe data.
        for i in range(netz_bunkers_1_rows):
            netz_marine_line.add_series({
                'name':       ['Bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + i + 7, 0],
                'categories': ['Bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + 6, 2,\
                    (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + 6, netz_bunkers_1_cols - 1],
                'values':     ['Bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + i + 7, netz_bunkers_1_cols - 1],
                'line':       {'color': netz_bunkers_1['fuel_code'].map(colours_dict).loc[i], 
                            'width': 1},
            })    
            
        ref_worksheet14.insert_chart('B' + str(chart_height + ref_bunkers_1_rows + ref_bunkers_2_rows + 9), netz_marine_line)

    else:
        pass

    # AVIATION BUNKER: Create a line chart subset by fuel
    if netz_bunkers_2_rows > 0:
        netz_aviation_line = workbook.add_chart({'type': 'line'})
        netz_aviation_line.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_aviation_line.set_chartarea({
            'border': {'none': True}
        })
        
        netz_aviation_line.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_aviation_line.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Aviation bunkers (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_aviation_line.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_aviation_line.set_title({
            'none': True
        }) 

        # Configure the series of the chart from the dataframe data.
        for i in range(netz_bunkers_2_rows):
            netz_aviation_line.add_series({
                'name':       ['Bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + i + 10, 0],
                'categories': ['Bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + 9, 2,\
                    (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + 9, netz_bunkers_2_cols - 1],
                'values':     ['Bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + i + 10, netz_bunkers_2_cols - 1],
                'line':       {'color': netz_bunkers_2['fuel_code'].map(colours_dict).loc[i], 
                            'width': 1},
            })    
            
        ref_worksheet14.insert_chart('J' + str(chart_height + ref_bunkers_1_rows + ref_bunkers_2_rows + 9), netz_aviation_line)

    else:
        pass

    #########################################################################################################################

    # TRANSFORMATION CHARTS

    # Access the workbook and first sheet with data from df1 
    ref_worksheet21 = writer.sheets['Power fuel consumption']
    
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
    ref_worksheet21.write(0, 0, economy + ' power input fuel Reference (NOTE: THIS IS NOT ELECTRICITY GENERATION)', cell_format1)
    ref_worksheet21.write(chart_height + ref_pow_use_2_rows + ref_pow_use_3_rows + 6, 0,\
        economy + ' power input fuel Carbon Neutrality (NOTE: THIS IS NOT ELECTRICITY GENERATION)', cell_format1)
    ref_worksheet21.write(1, 0, 'Units: Petajoules', cell_format2)

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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        usefuel_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        usefuel_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        usefuel_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_pow_use_2_rows):
            if not ref_pow_use_2['FUEL'].iloc[i] in ['Total']:
                usefuel_chart1.add_series({
                    'name':       ['Power fuel consumption', chart_height + i + 1, 0],
                    'categories': ['Power fuel consumption', chart_height, 2, chart_height, ref_pow_use_2_cols - 1],
                    'values':     ['Power fuel consumption', chart_height + i + 1, 2, chart_height + i + 1, ref_pow_use_2_cols - 1],
                    'fill':       {'color': ref_pow_use_2['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        usefuel_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        usefuel_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        usefuel_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_pow_use_3_rows):
            if not ref_pow_use_3['FUEL'].iloc[i] in ['Total']:
                usefuel_chart2.add_series({
                    'name':       ['Power fuel consumption', chart_height + ref_pow_use_2_rows + i + 4, 0],
                    'categories': ['Power fuel consumption', chart_height + ref_pow_use_2_rows + 3, 2, chart_height + ref_pow_use_2_rows + 3, ref_pow_use_3_cols - 1],
                    'values':     ['Power fuel consumption', chart_height + ref_pow_use_2_rows + i + 4, 2, chart_height + ref_pow_use_2_rows + i + 4, ref_pow_use_3_cols - 1],
                    'fill':       {'color': ref_pow_use_3['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass

        ref_worksheet21.insert_chart('J3', usefuel_chart2)

    else:
        pass

    # Create a use by fuel area chart
    if ref_pow_use_2_rows > 0:
        usefuel_chart3 = workbook.add_chart({'type': 'line'})
        usefuel_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        usefuel_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        usefuel_chart3.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        usefuel_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        usefuel_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        usefuel_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_pow_use_2_rows):
            if not ref_pow_use_2['FUEL'].iloc[i] in ['Total']:
                usefuel_chart3.add_series({
                    'name':       ['Power fuel consumption', chart_height + i + 1, 0],
                    'categories': ['Power fuel consumption', chart_height, 2, chart_height, ref_pow_use_2_cols - 1],
                    'values':     ['Power fuel consumption', chart_height + i + 1, 2, chart_height + i + 1, ref_pow_use_2_cols - 1],
                    'line':       {'color': ref_pow_use_2['FUEL'].map(colours_dict).loc[i], 'width': 1}
                })

            else:
                pass    
            
        ref_worksheet21.insert_chart('R3', usefuel_chart3)

    else:
        pass

    ############################# Next sheet: Production of electricity by technology ##################################
    
    # Access the workbook and second sheet
    ref_worksheet22 = writer.sheets['Generation']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet22.set_column(2, ref_elecgen_2_cols + 1, None, space_format)
    ref_worksheet22.set_row(chart_height, None, header_format)
    ref_worksheet22.set_row(chart_height + ref_elecgen_2_rows + 3, None, header_format)
    ref_worksheet22.set_row((2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6, None, header_format)
    ref_worksheet22.set_row((2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9, None, header_format)
    ref_worksheet22.write(0, 0, economy + ' electricity generation Reference', cell_format1)
    ref_worksheet22.write(chart_height + ref_elecgen_2_rows + ref_elecgen_3_rows + 6, 0, economy + ' electricity generation Carbon Neutrality', cell_format1)
    ref_worksheet22.write(1, 0, 'Units: Terrawatt hours', cell_format2)

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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        prodelec_bytech_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'TWh',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        prodelec_bytech_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        prodelec_bytech_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_elecgen_2_rows):
            if not ref_elecgen_2['TECHNOLOGY'].iloc[i] in ['Coal CCS', 'Gas CCS', 'Total']:
                prodelec_bytech_chart1.add_series({
                    'name':       ['Generation', chart_height + i + 1, 0],
                    'categories': ['Generation', chart_height, 2, chart_height, ref_elecgen_2_cols - 1],
                    'values':     ['Generation', chart_height + i + 1, 2, chart_height + i + 1, ref_elecgen_2_cols - 1],
                    'fill':       {'color': ref_elecgen_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                if not ref_elecgen_2['TECHNOLOGY'].iloc[i] in ['Total']:
                    prodelec_bytech_chart1.add_series({
                        'name':       ['Generation', chart_height + i + 1, 0],
                        'categories': ['Generation', chart_height, 2, chart_height, ref_elecgen_2_cols - 1],
                        'values':     ['Generation', chart_height + i + 1, 2, chart_height + i + 1, ref_elecgen_2_cols - 1],
                        'pattern':    {'fg_color': ref_elecgen_2['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True}
                    })

                else:
                    pass
            
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        prodelec_bytech_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'TWh',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        prodelec_bytech_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        prodelec_bytech_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_elecgen_3_rows):
            if not ref_elecgen_3['TECHNOLOGY'].iloc[i] in ['Coal CCS', 'Gas CCS', 'Total']:
                prodelec_bytech_chart2.add_series({
                    'name':       ['Generation', chart_height + ref_elecgen_2_rows + i + 4, 0],
                    'categories': ['Generation', chart_height + ref_elecgen_2_rows + 3, 2, chart_height + ref_elecgen_2_rows + 3, ref_elecgen_3_cols - 1],
                    'values':     ['Generation', chart_height + ref_elecgen_2_rows + i + 4, 2, chart_height + ref_elecgen_2_rows + i + 4, ref_elecgen_3_cols - 1],
                    'fill':       {'color': ref_elecgen_3['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                if not ref_elecgen_3['TECHNOLOGY'].iloc[i] in ['Total']:
                    prodelec_bytech_chart2.add_series({
                        'name':       ['Generation', chart_height + ref_elecgen_2_rows + i + 4, 0],
                        'categories': ['Generation', chart_height + ref_elecgen_2_rows + 3, 2, chart_height + ref_elecgen_2_rows + 3, ref_elecgen_3_cols - 1],
                        'values':     ['Generation', chart_height + ref_elecgen_2_rows + i + 4, 2, chart_height + ref_elecgen_2_rows + i + 4, ref_elecgen_3_cols - 1],
                        'pattern':    {'fg_color': ref_elecgen_3['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True},
                        'gap':        100
                    })
                
                else:
                    pass
            
        ref_worksheet22.insert_chart('J3', prodelec_bytech_chart2)
    
    else:
        pass

    #################################################################################################################################################

    ## Refining sheet

    # Access the workbook and second sheet
    ref_worksheet23 = writer.sheets['Refining']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet23.set_column(2, ref_refinery_1_cols + 1, None, space_format)
    ref_worksheet23.set_row(chart_height, None, header_format)
    ref_worksheet23.set_row(chart_height + ref_refinery_1_rows + 3, None, header_format)
    ref_worksheet23.set_row(chart_height + ref_refinery_1_rows + ref_refinery_2_rows + 6, None, header_format)
    ref_worksheet23.set_row((2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9, None, header_format)
    ref_worksheet23.set_row((2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + 12, None, header_format)
    ref_worksheet23.set_row((2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + 15, None, header_format)
    ref_worksheet23.write(0, 0, economy + ' refining Reference', cell_format1)
    ref_worksheet23.write(chart_height + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9, 0, economy + ' refining Carbon Neutrality', cell_format1)
    ref_worksheet23.write(1, 0, 'Units: Petajoules', cell_format2)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        refinery_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        refinery_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        refinery_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_refinery_1_rows):
            if not ref_refinery_1['FUEL'].iloc[i] in ['Total']:
                refinery_chart1.add_series({
                    'name':       ['Refining', chart_height + i + 1, 0],
                    'categories': ['Refining', chart_height, 2, chart_height, ref_refinery_1_cols - 1],
                    'values':     ['Refining', chart_height + i + 1, 2, chart_height + i + 1, ref_refinery_1_cols - 1],
                    'line':       {'color': ref_refinery_1['FUEL'].map(colours_dict).loc[i],
                                'width': 1}
                })

            else:
                pass    
            
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        refinery_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        refinery_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        refinery_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_refinery_2_rows):
            if not ref_refinery_2['FUEL'].iloc[i] in ['Total']:
                refinery_chart2.add_series({
                    'name':       ['Refining', chart_height + ref_refinery_1_rows + i + 4, 0],
                    'categories': ['Refining', chart_height + ref_refinery_1_rows + 3, 2, chart_height + ref_refinery_1_rows + 3, ref_refinery_2_cols - 1],
                    'values':     ['Refining', chart_height + ref_refinery_1_rows + i + 4, 2, chart_height + ref_refinery_1_rows + i + 4, ref_refinery_2_cols - 1],
                    'line':       {'color': ref_refinery_2['FUEL'].map(colours_dict).loc[i],
                                'width': 1}
                })

            else:
                pass    
            
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        refinery_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        refinery_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        refinery_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_refinery_3_rows):
            if not ref_refinery_3['FUEL'].iloc[i] in ['Total']:
                refinery_chart3.add_series({
                    'name':       ['Refining', chart_height + ref_refinery_1_rows + ref_refinery_2_rows + i + 7, 0],
                    'categories': ['Refining', chart_height + ref_refinery_1_rows + ref_refinery_2_rows + 6, 2, chart_height + ref_refinery_1_rows + ref_refinery_2_rows + 6, ref_refinery_3_cols - 1],
                    'values':     ['Refining', chart_height + ref_refinery_1_rows + ref_refinery_2_rows + i + 7, 2, chart_height + ref_refinery_1_rows + ref_refinery_2_rows + i + 7, ref_refinery_3_cols - 1],
                    'fill':       {'color': ref_refinery_3['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass    
            
        ref_worksheet23.insert_chart('R3', refinery_chart3)

    else:
        pass

    ############################# Next sheet: Power capacity ##################################
    
    # Access the workbook and second sheet
    ref_worksheet24 = writer.sheets['Capacity']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet24.set_column(1, ref_powcap_1_cols + 1, None, space_format)
    ref_worksheet24.set_row(chart_height, None, header_format)
    ref_worksheet24.set_row(chart_height + ref_powcap_1_rows + 3, None, header_format)
    ref_worksheet24.set_row((2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6, None, header_format)
    ref_worksheet24.set_row((2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9, None, header_format)
    ref_worksheet24.write(0, 0, economy + ' power capacity Reference', cell_format1)
    ref_worksheet24.write(chart_height + ref_powcap_1_rows + ref_powcap_2_rows + 6, 0, economy + ' power capacity Carbon Neutrality', cell_format1)
    ref_worksheet24.write(1, 0, 'Units: Gigawatts', cell_format2)

    # Create a electricity production area chart
    if ref_powcap_1_rows > 1:
        pow_cap_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        pow_cap_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        pow_cap_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        pow_cap_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        pow_cap_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'GW',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        pow_cap_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        pow_cap_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_powcap_1_rows):
            if not ref_powcap_1['TECHNOLOGY'].iloc[i] in ['Coal CCS', 'Gas CCS', 'Total']:
                pow_cap_chart1.add_series({
                    'name':       ['Capacity', chart_height + i + 1, 0],
                    'categories': ['Capacity', chart_height, 1, chart_height, ref_powcap_1_cols - 1],
                    'values':     ['Capacity', chart_height + i + 1, 1, chart_height + i + 1, ref_powcap_1_cols - 1],
                    'fill':       {'color': ref_powcap_1['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                if not ref_powcap_1['TECHNOLOGY'].iloc[i] in ['Total']:
                    pow_cap_chart1.add_series({
                        'name':       ['Capacity', chart_height + i + 1, 0],
                        'categories': ['Capacity', chart_height, 1, chart_height, ref_powcap_1_cols - 1],
                        'values':     ['Capacity', chart_height + i + 1, 1, chart_height + i + 1, ref_powcap_1_cols - 1],
                        'pattern':    {'fg_color': ref_powcap_1['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True}
                    })

                else:
                    pass    
            
        ref_worksheet24.insert_chart('B3', pow_cap_chart1)

    else:
        pass

    # Create a industry subsector FED chart
    if ref_powcap_2_rows > 1:
        pow_cap_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        pow_cap_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        pow_cap_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        pow_cap_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        pow_cap_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'GW',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        pow_cap_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        pow_cap_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_powcap_2_rows):
            if not ref_powcap_2['TECHNOLOGY'].iloc[i] in ['Coal CCS', 'Gas CCS', 'Total']:
                pow_cap_chart2.add_series({
                    'name':       ['Capacity', chart_height + ref_powcap_1_rows + i + 4, 0],
                    'categories': ['Capacity', chart_height + ref_powcap_1_rows + 3, 1, chart_height + ref_powcap_1_rows + 3, ref_powcap_2_cols - 1],
                    'values':     ['Capacity', chart_height + ref_powcap_1_rows + i + 4, 1, chart_height + ref_powcap_1_rows + i + 4, ref_powcap_2_cols - 1],
                    'fill':       {'color': ref_powcap_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                if not ref_powcap_2['TECHNOLOGY'].iloc[i] in ['Total']:
                    pow_cap_chart2.add_series({
                        'name':       ['Capacity', chart_height + ref_powcap_1_rows + i + 4, 0],
                        'categories': ['Capacity', chart_height + ref_powcap_1_rows + 3, 1, chart_height + ref_powcap_1_rows + 3, ref_powcap_2_cols - 1],
                        'values':     ['Capacity', chart_height + ref_powcap_1_rows + i + 4, 1, chart_height + ref_powcap_1_rows + i + 4, ref_powcap_2_cols - 1],
                        'pattern':    {'fg_color': ref_powcap_2['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True},
                        'gap':        100
                    })

                else:
                    pass    
            
        ref_worksheet24.insert_chart('J3', pow_cap_chart2)

    else:
        pass

    ############################# Next sheet: Transformation sector ##################################
    
    # Access the workbook and second sheet
    ref_worksheet25 = writer.sheets['Transformation']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet25.set_column(1, ref_trans_3_cols + 1, None, space_format)
    ref_worksheet25.set_row(chart_height, None, header_format)
    ref_worksheet25.set_row(chart_height + ref_trans_3_rows + 3, None, header_format)
    ref_worksheet25.set_row((2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, None, header_format)
    ref_worksheet25.set_row((2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + ref_trans_3_rows + 9, None, header_format)
    ref_worksheet25.write(0, 0, economy + ' transformation Reference', cell_format1)
    ref_worksheet25.write(chart_height + ref_trans_3_rows + ref_trans_4_rows + 6, 0, economy + ' transformation Carbon Neutrality', cell_format1)
    ref_worksheet25.write(1, 0, 'Units: Petajoules', cell_format2)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_trnsfrm_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_trnsfrm_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_trnsfrm_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_trans_3_rows):
            ref_trnsfrm_chart1.add_series({
                'name':       ['Transformation', chart_height + i + 1, 0],
                'categories': ['Transformation', chart_height, 1, chart_height, ref_trans_3_cols - 1],
                'values':     ['Transformation', chart_height + i + 1, 1, chart_height + i + 1, ref_trans_3_cols - 1],
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_trnsfrm_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_trnsfrm_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_trnsfrm_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_trans_3_rows):
            ref_trnsfrm_chart2.add_series({
                'name':       ['Transformation', chart_height + i + 1, 0],
                'categories': ['Transformation', chart_height, 1, chart_height, ref_trans_3_cols - 1],
                'values':     ['Transformation', chart_height + i + 1, 1, chart_height + i + 1, ref_trans_3_cols - 1],
                'line':       {'color': ref_trans_3['Sector'].map(colours_dict).loc[i],
                               'width': 1}
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_trnsfrm_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_trnsfrm_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_trnsfrm_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_trans_4_rows):
            ref_trnsfrm_chart3.add_series({
                'name':       ['Transformation', chart_height + ref_trans_3_rows + i + 4, 0],
                'categories': ['Transformation', chart_height + ref_trans_3_rows + 3, 1, chart_height + ref_trans_3_rows + 3, ref_trans_4_cols - 1],
                'values':     ['Transformation', chart_height + ref_trans_3_rows + i + 4, 1, chart_height + ref_trans_3_rows + i + 4, ref_trans_4_cols - 1],
                'fill':       {'color': ref_trans_4['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })    
            
        ref_worksheet25.insert_chart('R3', ref_trnsfrm_chart3)

    else:
        pass

    ###############################################################################
    # Own use charts
    
    # Access the workbook and second sheet
    ref_worksheet26 = writer.sheets['Own-use']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet26.set_column(2, ref_ownuse_1_cols + 1, None, space_format)
    ref_worksheet26.set_row(chart_height, None, header_format)
    ref_worksheet26.set_row(chart_height + ref_ownuse_1_rows + 3, None, header_format)
    ref_worksheet26.set_row((2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + 6, None, header_format)
    ref_worksheet26.set_row((2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + 9, None, header_format)
    ref_worksheet26.write(0, 0, economy + ' own use and losses Reference', cell_format1)
    ref_worksheet26.write(chart_height + ref_ownuse_1_rows + ref_ownuse_2_rows + 6, 0, economy + ' own use and losses Carbon Neutrality', cell_format1)
    ref_worksheet26.write(1, 0, 'Units: Petajoules', cell_format2)

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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_own_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_own_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_own_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_ownuse_1_rows):
            if not ref_ownuse_1['FUEL'].iloc[i] in ['Total']:
                ref_own_chart1.add_series({
                    'name':       ['Own-use', chart_height + i + 1, 0],
                    'categories': ['Own-use', chart_height, 2, chart_height, ref_ownuse_1_cols - 1],
                    'values':     ['Own-use', chart_height + i + 1, 2, chart_height + i + 1, ref_ownuse_1_cols - 1],
                    'fill':       {'color': ref_ownuse_1['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_own_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_own_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_own_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_ownuse_1_rows):
            if not ref_ownuse_1['FUEL'].iloc[i] in ['Total']:
                ref_own_chart2.add_series({
                    'name':       ['Own-use', chart_height + i + 1, 0],
                    'categories': ['Own-use', chart_height, 2, chart_height, ref_ownuse_1_cols - 1],
                    'values':     ['Own-use', chart_height + i + 1, 2, chart_height + i + 1, ref_ownuse_1_cols - 1],
                    'line':       {'color': ref_ownuse_1['FUEL'].map(colours_dict).loc[i],
                                'width': 1}
                })

            else:
                pass    
            
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_own_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_own_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_own_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_ownuse_2_rows):
            if not ref_ownuse_2['FUEL'].iloc[i] in ['Total']:
                ref_own_chart3.add_series({
                    'name':       ['Own-use', chart_height + ref_ownuse_1_rows + i + 4, 0],
                    'categories': ['Own-use', chart_height + ref_ownuse_1_rows + 3, 2, chart_height + ref_ownuse_1_rows + 3, ref_ownuse_2_cols - 1],
                    'values':     ['Own-use', chart_height + ref_ownuse_1_rows + i + 4, 2, chart_height + ref_ownuse_1_rows + i + 4, ref_ownuse_2_cols - 1],
                    'fill':       {'color': ref_ownuse_2['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass    
            
        ref_worksheet26.insert_chart('R3', ref_own_chart3)

    else:
        pass

    ############## HEAT Charts #########################################

    # Access the workbook and second sheet
    ref_worksheet27 = writer.sheets['Heat generation']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet27.set_column(2, ref_heatgen_2_cols + 1, None, space_format)
    ref_worksheet27.set_row(chart_height, None, header_format)
    ref_worksheet27.set_row(chart_height + ref_heatgen_2_rows + 3, None, header_format)
    ref_worksheet27.set_row((2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6, None, header_format)
    ref_worksheet27.set_row((2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9, None, header_format)
    ref_worksheet27.write(0, 0, economy + ' heat generation Reference (note: the category "heat only units" consume multiple different fuels, depending on the economy)', cell_format1)
    ref_worksheet27.write(chart_height + ref_heatgen_2_rows + ref_heatgen_3_rows + 6, 0, economy +\
        ' heat generation Carbon Neutrality (note: the category "heat only units" consume multiple different fuels, depending on the economy)', cell_format1)
    ref_worksheet27.write(1, 0, 'Units: Petajoules', cell_format2)

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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        heatgen_bytech_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        heatgen_bytech_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_heatgen_2_rows):
            if not ref_heatgen_2['TECHNOLOGY'].iloc[i] in ['Gas CCS', 'Total']:
                heatgen_bytech_chart1.add_series({
                    'name':       ['Heat generation', chart_height + i + 1, 0],
                    'categories': ['Heat generation', chart_height, 2, chart_height, ref_heatgen_2_cols - 1],
                    'values':     ['Heat generation', chart_height + i + 1, 2, chart_height + i + 1, ref_heatgen_2_cols - 1],
                    'fill':       {'color': ref_heatgen_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                if not ref_heatgen_2['TECHNOLOGY'].iloc[i] in ['Total']:
                    heatgen_bytech_chart1.add_series({
                        'name':       ['Heat generation', chart_height + i + 1, 0],
                        'categories': ['Heat generation', chart_height, 2, chart_height, ref_heatgen_2_cols - 1],
                        'values':     ['Heat generation', chart_height + i + 1, 2, chart_height + i + 1, ref_heatgen_2_cols - 1],
                        'pattern':    {'fg_color': ref_heatgen_2['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True}
                    })

                else:
                    pass    
            
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        heatgen_bytech_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_heatgen_3_rows):
            if not ref_heatgen_3['TECHNOLOGY'].iloc[i] in ['Gas CCS', 'Total']:
                heatgen_bytech_chart2.add_series({
                    'name':       ['Heat generation', chart_height + ref_heatgen_2_rows + i + 4, 0],
                    'categories': ['Heat generation', chart_height + ref_heatgen_2_rows + 3, 2, chart_height + ref_heatgen_2_rows + 3, ref_heatgen_3_cols - 1],
                    'values':     ['Heat generation', chart_height + ref_heatgen_2_rows + i + 4, 2, chart_height + ref_heatgen_2_rows + i + 4, ref_heatgen_3_cols - 1],
                    'fill':       {'color': ref_heatgen_3['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                if not ref_heatgen_3['TECHNOLOGY'].iloc[i] in ['Total']:
                    heatgen_bytech_chart2.add_series({
                        'name':       ['Heat generation', chart_height + ref_heatgen_2_rows + i + 4, 0],
                        'categories': ['Heat generation', chart_height + ref_heatgen_2_rows + 3, 2, chart_height + ref_heatgen_2_rows + 3, ref_heatgen_3_cols - 1],
                        'values':     ['Heat generation', chart_height + ref_heatgen_2_rows + i + 4, 2, chart_height + ref_heatgen_2_rows + i + 4, ref_heatgen_3_cols - 1],
                        'pattern':    {'fg_color': ref_heatgen_3['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True},
                        'gap':        100
                    })

                else:
                    pass    
            
        ref_worksheet27.insert_chart('J3', heatgen_bytech_chart2)
    
    else:
        pass

    ##########################################################################

    # Access the workbook and first sheet with data from df1 
    ref_worksheet28 = writer.sheets['Heat only consumption']
    
    # Comma format and header format        
    # space_format = workbook.add_format({'num_format': '#,##0'})
    # header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    # cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet28.set_column(2, ref_heat_use_2_cols + 1, None, space_format)
    ref_worksheet28.set_row(chart_height, None, header_format)
    ref_worksheet28.set_row(chart_height + ref_heat_use_2_rows + 3, None, header_format)
    ref_worksheet28.set_row((2 * chart_height) + ref_heat_use_2_rows + ref_heat_use_3_rows + 6, None, header_format)
    ref_worksheet28.set_row((2 * chart_height) + ref_heat_use_2_rows + ref_heat_use_3_rows + netz_heat_use_2_rows + 9, None, header_format)
    ref_worksheet28.write(0, 0, economy + ' heat input fuel Reference', cell_format1)
    ref_worksheet28.write(chart_height + ref_heat_use_2_rows + ref_heat_use_3_rows + 6, 0,\
        economy + ' heat input fuel Carbon Neutrality', cell_format1)
    ref_worksheet28.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a use by fuel area chart
    if ref_heat_use_2_rows > 0:
        ref_heatuse_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_heatuse_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_heatuse_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_heatuse_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_heatuse_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_heatuse_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_heatuse_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_heat_use_2_rows):
            if not ref_heat_use_2['FUEL'].iloc[i] in ['Total']:
                ref_heatuse_chart1.add_series({
                    'name':       ['Heat only consumption', chart_height + i + 1, 0],
                    'categories': ['Heat only consumption', chart_height, 2, chart_height, ref_heat_use_2_cols - 1],
                    'values':     ['Heat only consumption', chart_height + i + 1, 2, chart_height + i + 1, ref_heat_use_2_cols - 1],
                    'fill':       {'color': ref_heat_use_2['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet28.insert_chart('B3', ref_heatuse_chart1)

    else:
        pass

    # Create a use column chart
    if ref_heat_use_3_rows > 0:
        ref_heatuse_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_heatuse_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_heatuse_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        ref_heatuse_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_heatuse_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_heatuse_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_heatuse_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_heat_use_3_rows):
            if not ref_heat_use_3['FUEL'].iloc[i] in ['Total']:
                ref_heatuse_chart2.add_series({
                    'name':       ['Heat only consumption', chart_height + ref_heat_use_2_rows + i + 4, 0],
                    'categories': ['Heat only consumption', chart_height + ref_heat_use_2_rows + 3, 2, chart_height + ref_heat_use_2_rows + 3, ref_heat_use_3_cols - 1],
                    'values':     ['Heat only consumption', chart_height + ref_heat_use_2_rows + i + 4, 2, chart_height + ref_heat_use_2_rows + i + 4, ref_heat_use_3_cols - 1],
                    'fill':       {'color': ref_heat_use_3['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass

        ref_worksheet28.insert_chart('J3', ref_heatuse_chart2)

    else:
        pass

    ########################################################################################################################################

    ################# NET ZERO CHARTS ######################################################################################################

    # Access the workbook and first sheet with data from df1
    # netz_worksheet21 = writer.sheets['Power consumption']
    
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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_usefuel_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_usefuel_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_usefuel_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_pow_use_2_rows):
            if not netz_pow_use_2['FUEL'].iloc[i] in ['Total']:
                netz_usefuel_chart1.add_series({
                    'name':       ['Power fuel consumption', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + i + 7, 0],
                    'categories': ['Power fuel consumption', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + 6, 2,\
                        (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + 6, netz_pow_use_2_cols - 1],
                    'values':     ['Power fuel consumption', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + i + 7, 2,\
                        (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + i + 7, netz_pow_use_2_cols - 1],
                    'fill':       {'color': netz_pow_use_2['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet21.insert_chart('B' + str(chart_height + ref_pow_use_2_rows + ref_pow_use_3_rows + 9), netz_usefuel_chart1)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_usefuel_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_usefuel_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_usefuel_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(netz_pow_use_3_rows):
            if not netz_pow_use_3['FUEL'].iloc[i] in ['Total']:
                netz_usefuel_chart2.add_series({
                    'name':       ['Power fuel consumption', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + i + 10, 0],
                    'categories': ['Power fuel consumption', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + 9, 2,\
                        (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + 9, netz_pow_use_3_cols - 1],
                    'values':     ['Power fuel consumption', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + i + 10, 2,\
                        (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + i + 10, netz_pow_use_3_cols - 1],
                    'fill':       {'color': netz_pow_use_3['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass

        ref_worksheet21.insert_chart('J' + str(chart_height + ref_pow_use_2_rows + ref_pow_use_3_rows + 9), netz_usefuel_chart2)

    else:
        pass
    
    # Line chart
    if netz_pow_use_2_rows > 0:
        netz_usefuel_chart3 = workbook.add_chart({'type': 'line'})
        netz_usefuel_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_usefuel_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        netz_usefuel_chart3.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_usefuel_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_usefuel_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_usefuel_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_pow_use_2_rows):
            if not netz_pow_use_2['FUEL'].iloc[i] in ['Total']:
                netz_usefuel_chart3.add_series({
                    'name':       ['Power fuel consumption', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + i + 7, 0],
                    'categories': ['Power fuel consumption', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + 6, 2,\
                        (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + 6, netz_pow_use_2_cols - 1],
                    'values':     ['Power fuel consumption', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + i + 7, 2,\
                        (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + i + 7, netz_pow_use_2_cols - 1],
                    'line':       {'color': netz_pow_use_2['FUEL'].map(colours_dict).loc[i], 'width': 1}
                })

            else:
                pass    
            
        ref_worksheet21.insert_chart('R' + str(chart_height + ref_pow_use_2_rows + ref_pow_use_3_rows + 9), netz_usefuel_chart3)

    else:
        pass

    ############################# Next sheet: Production of electricity by technology ##################################
    
    # Access the workbook and second sheet
    # netz_worksheet22 = writer.sheets['Generation']
    
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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_prodelec_bytech_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'TWh',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_prodelec_bytech_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_prodelec_bytech_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_elecgen_2_rows):
            if not netz_elecgen_2['TECHNOLOGY'].iloc[i] in ['Coal CCS', 'Gas CCS', 'Total']:    
                netz_prodelec_bytech_chart1.add_series({
                    'name':       ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + i + 7, 0],
                    'categories': ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6, 2,\
                        (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6, netz_elecgen_2_cols - 1],
                    'values':     ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + i + 7, 2,\
                        (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + i + 7, netz_elecgen_2_cols - 1],
                    'fill':       {'color': netz_elecgen_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                if not netz_elecgen_2['TECHNOLOGY'].iloc[i] in ['Total']:
                    netz_prodelec_bytech_chart1.add_series({
                        'name':       ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + i + 7, 0],
                        'categories': ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6, 2,\
                            (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6, netz_elecgen_2_cols - 1],
                        'values':     ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + i + 7, 2,\
                            (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + i + 7, netz_elecgen_2_cols - 1],
                        'pattern':    {'fg_color': netz_elecgen_2['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True}
                    })

                else:
                    pass
           
        ref_worksheet22.insert_chart('B' + str(chart_height + ref_elecgen_2_rows + ref_elecgen_3_rows + 9), netz_prodelec_bytech_chart1)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_prodelec_bytech_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'TWh',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_prodelec_bytech_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_prodelec_bytech_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_elecgen_3_rows):
            if not netz_elecgen_3['TECHNOLOGY'].iloc[i] in ['Coal CCS', 'Gas CCS', 'Total']:
                netz_prodelec_bytech_chart2.add_series({
                    'name':       ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + i + 10, 0],
                    'categories': ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9, 2,\
                        (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9, netz_elecgen_3_cols - 1],
                    'values':     ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + i + 10, 2,\
                        (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + i + 10, netz_elecgen_3_cols - 1],
                    'fill':       {'color': netz_elecgen_3['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                if not netz_elecgen_3['TECHNOLOGY'].iloc[i] in ['Total']: 
                    netz_prodelec_bytech_chart2.add_series({
                        'name':       ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + i + 10, 0],
                        'categories': ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9, 2,\
                            (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9, netz_elecgen_3_cols - 1],
                        'values':     ['Generation', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + i + 10, 2,\
                            (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + i + 10, netz_elecgen_3_cols - 1],
                        'pattern':    {'fg_color': netz_elecgen_3['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True},
                        'gap':        100
                    })   
            
        ref_worksheet22.insert_chart('J' + str(chart_height + ref_elecgen_2_rows + ref_elecgen_3_rows + 9), netz_prodelec_bytech_chart2)
    
    else:
        pass

    #################################################################################################################################################

    ## Refining sheet

    # Access the workbook and second sheet
    # netz_worksheet23 = writer.sheets['Refining']
    
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_refinery_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_refinery_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_refinery_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_refinery_1_rows):
            if not netz_refinery_1['FUEL'].iloc[i] in ['Total']:
                netz_refinery_chart1.add_series({
                    'name':       ['Refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + i + 10, 0],
                    'categories': ['Refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9, 2,\
                        (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9, netz_refinery_1_cols - 1],
                    'values':     ['Refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + i + 10, 2,\
                        (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + i + 10, netz_refinery_1_cols - 1],
                    'line':       {'color': netz_refinery_1['FUEL'].map(colours_dict).loc[i],
                                'width': 1}
                })

            else:
                pass    
            
        ref_worksheet23.insert_chart('B' + str(chart_height + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 12), netz_refinery_chart1)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_refinery_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_refinery_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_refinery_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_refinery_2_rows):
            if not netz_refinery_2['FUEL'].iloc[i] in ['Total']:
                netz_refinery_chart2.add_series({
                    'name':       ['Refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + i + 13, 0],
                    'categories': ['Refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + 12, 2,\
                        (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + 12, netz_refinery_2_cols - 1],
                    'values':     ['Refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + i + 13, 2,\
                        (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + i + 13, netz_refinery_2_cols - 1],
                    'line':       {'color': netz_refinery_2['FUEL'].map(colours_dict).loc[i],
                                'width': 1}
                })

            else:
                pass    
            
        ref_worksheet23.insert_chart('J' + str(chart_height + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 12), netz_refinery_chart2)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_refinery_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_refinery_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_refinery_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_refinery_3_rows):
            if not netz_refinery_3['FUEL'].iloc[i] in ['Total']:
                netz_refinery_chart3.add_series({
                    'name':       ['Refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + i + 16, 0],
                    'categories': ['Refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + 15, 2,\
                        (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + 15, netz_refinery_3_cols - 1],
                    'values':     ['Refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + i + 16, 2,\
                        (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + i + 16, netz_refinery_3_cols - 1],
                    'fill':       {'color': netz_refinery_3['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass    
            
        ref_worksheet23.insert_chart('R' + str(chart_height + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 12), netz_refinery_chart3)

    else:
        pass

    ############################# Next sheet: Power capacity ##################################
    
    # Access the workbook and second sheet
    # netz_worksheet24 = writer.sheets['Capacity']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet24.set_column(1, netz_powcap_1_cols + 1, None, space_format)
    # netz_worksheet24.set_row(chart_height, None, header_format)
    # netz_worksheet24.set_row(chart_height + netz_powcap_1_rows + 3, None, header_format)
    # netz_worksheet24.write(0, 0, economy + ' electricity capacity by technology', cell_format1)
    
    # Create a electricity production area chart
    if netz_powcap_1_rows > 1:
        netz_pow_cap_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_pow_cap_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_pow_cap_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_pow_cap_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_pow_cap_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'GW',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_pow_cap_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_pow_cap_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_powcap_1_rows):
            if not netz_powcap_1['TECHNOLOGY'].iloc[i] in ['Coal CCS', 'Gas CCS', 'Total']:
                netz_pow_cap_chart1.add_series({
                    'name':       ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + i + 7, 0],
                    'categories': ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6, 1,\
                        (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6, netz_powcap_1_cols - 1],
                    'values':     ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + i + 7, 1,\
                        (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + i + 7, netz_powcap_1_cols - 1],
                    'fill':       {'color': netz_powcap_1['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                if not netz_powcap_1['TECHNOLOGY'].iloc[i] in ['Total']:
                    netz_pow_cap_chart1.add_series({
                        'name':       ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + i + 7, 0],
                        'categories': ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6, 1,\
                            (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6, netz_powcap_1_cols - 1],
                        'values':     ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + i + 7, 1,\
                            (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + i + 7, netz_powcap_1_cols - 1],
                        'pattern':    {'fg_color': netz_powcap_1['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True}
                    })

                else:
                    pass    

        ref_worksheet24.insert_chart('B' + str(chart_height + ref_powcap_1_rows + ref_powcap_2_rows + 9), netz_pow_cap_chart1)

    else:
        pass

    # Create a industry subsector FED chart
    if netz_powcap_2_rows > 1:
        netz_pow_cap_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_pow_cap_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_pow_cap_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_pow_cap_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_pow_cap_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'GW',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_pow_cap_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_pow_cap_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_powcap_2_rows):
            if not netz_powcap_2['TECHNOLOGY'].iloc[i] in ['Coal CCS', 'Gas CCS', 'Total']:
                netz_pow_cap_chart2.add_series({
                    'name':       ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + i + 10, 0],
                    'categories': ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9, 1,\
                        (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9, netz_powcap_2_cols - 1],
                    'values':     ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + i + 10, 1,\
                        (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + i + 10, netz_powcap_2_cols - 1],
                    'fill':       {'color': netz_powcap_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                if not netz_powcap_2['TECHNOLOGY'].iloc[i] in ['Total']:
                    netz_pow_cap_chart2.add_series({
                        'name':       ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + i + 10, 0],
                        'categories': ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9, 1,\
                            (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9, netz_powcap_2_cols - 1],
                        'values':     ['Capacity', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + i + 10, 1,\
                            (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + i + 10, netz_powcap_2_cols - 1],
                        'pattern':    {'fg_color': netz_powcap_2['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True},
                        'gap':        100
                    })

                else:
                    pass    
            
        ref_worksheet24.insert_chart('J' + str(chart_height + ref_powcap_1_rows + ref_powcap_2_rows + 9), netz_pow_cap_chart2)

    else:
        pass

    ############################# Next sheet: Transformation sector ##################################
    
    # Access the workbook and second sheet
    # netz_worksheet25 = writer.sheets['Transformation']
    
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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_trnsfrm_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_trnsfrm_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_trnsfrm_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_trans_3_rows):
            netz_trnsfrm_chart1.add_series({
                'name':       ['Transformation', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, 0],
                'categories': ['Transformation', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, netz_trans_3_cols - 1],
                'values':     ['Transformation', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, netz_trans_3_cols - 1],
                'fill':       {'color': netz_trans_3['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet25.insert_chart('B' + str(chart_height + ref_trans_3_rows + ref_trans_4_rows + 9), netz_trnsfrm_chart1)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_trnsfrm_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_trnsfrm_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_trnsfrm_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_trans_3_rows):
            netz_trnsfrm_chart2.add_series({
                'name':       ['Transformation', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, 0],
                'categories': ['Transformation', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, netz_trans_3_cols - 1],
                'values':     ['Transformation', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, netz_trans_3_cols - 1],
                'line':       {'color': netz_trans_3['Sector'].map(colours_dict).loc[i],
                               'width': 1}
            })    
            
        ref_worksheet25.insert_chart('J' + str(chart_height + ref_trans_3_rows + ref_trans_4_rows + 9), netz_trnsfrm_chart2)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_trnsfrm_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_trnsfrm_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_trnsfrm_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_trans_4_rows):
            netz_trnsfrm_chart3.add_series({
                'name':       ['Transformation', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + i + 10, 0],
                'categories': ['Transformation', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + 9, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + 9, netz_trans_4_cols - 1],
                'values':     ['Transformation', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + i + 10, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + i + 10, netz_trans_4_cols - 1],
                'fill':       {'color': netz_trans_4['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })    
            
        ref_worksheet25.insert_chart('R' + str(chart_height + ref_trans_3_rows + ref_trans_4_rows + 9), netz_trnsfrm_chart3)

    else:
        pass

    ###############################################################################
    # Own use charts
    
    # Access the workbook and second sheet
    # netz_worksheet26 = writer.sheets['Own-use']
    
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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_own_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_own_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_own_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_ownuse_1_rows):
            if not netz_ownuse_1['FUEL'].iloc[i] in ['Total']:
                netz_own_chart1.add_series({
                    'name':       ['Own-use', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + i + 7, 0],
                    'categories': ['Own-use', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + 6, 2,\
                        (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + 6, netz_ownuse_1_cols - 1],
                    'values':     ['Own-use', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + i + 7, netz_ownuse_1_cols - 1],
                    'fill':       {'color': netz_ownuse_1['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet26.insert_chart('B' + str(chart_height + ref_ownuse_1_rows + ref_ownuse_2_rows + 9), netz_own_chart1)

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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_own_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_own_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_own_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_ownuse_1_rows):
            if not netz_ownuse_1['FUEL'].iloc[i] in ['Total']:
                netz_own_chart2.add_series({
                    'name':       ['Own-use', chart_height + i + 1, 0],
                    'categories': ['Own-use', chart_height, 2, chart_height, netz_ownuse_1_cols - 1],
                    'values':     ['Own-use', chart_height + i + 1, 2, chart_height + i + 1, netz_ownuse_1_cols - 1],
                    'line':       {'color': netz_ownuse_1['FUEL'].map(colours_dict).loc[i],
                                'width': 1}
                })

            else:
                pass    
            
        ref_worksheet26.insert_chart('J' + str(chart_height + ref_ownuse_1_rows + ref_ownuse_2_rows + 9), netz_own_chart2)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_own_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_own_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_own_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_ownuse_2_rows):
            if not netz_ownuse_2['FUEL'].iloc[i] in ['Total']:
                netz_own_chart3.add_series({
                    'name':       ['Own-use', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + i + 10, 0],
                    'categories': ['Own-use', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + 9, 2,\
                        (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + 9, netz_ownuse_2_cols - 1],
                    'values':     ['Own-use', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + i + 10, 2,\
                        (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + i + 10, netz_ownuse_2_cols - 1],
                    'fill':       {'color': netz_ownuse_2['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass    
            
        ref_worksheet26.insert_chart('R' + str(chart_height + ref_ownuse_1_rows + ref_ownuse_2_rows + 9), netz_own_chart3)

    else:
        pass

    ##################### HEAT Charts ###########################################

    # Access the workbook and second sheet
    # netz_worksheet27 = writer.sheets['Heat generation']
    
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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        heatgen_bytech_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        heatgen_bytech_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_heatgen_2_rows):
            if not netz_heatgen_2['TECHNOLOGY'].iloc[i] in ['Gas CCS', 'Total']: 
                heatgen_bytech_chart1.add_series({
                    'name':       ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + i + 7, 0],
                    'categories': ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6, 2,\
                        (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6, netz_heatgen_2_cols - 1],
                    'values':     ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + i + 7, 2,\
                        (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + i + 7, netz_heatgen_2_cols - 1],
                    'fill':       {'color': netz_heatgen_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                if not netz_heatgen_2['TECHNOLOGY'].iloc[i] in ['Total']:
                    heatgen_bytech_chart1.add_series({
                        'name':       ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + i + 7, 0],
                        'categories': ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6, 2,\
                            (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6, netz_heatgen_2_cols - 1],
                        'values':     ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + i + 7, 2,\
                            (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + i + 7, netz_heatgen_2_cols - 1],
                        'pattern':    {'fg_color': netz_heatgen_2['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True}
                    })

                else:
                    pass    
            
        ref_worksheet27.insert_chart('B' + str(chart_height + ref_heatgen_2_rows + ref_heatgen_3_rows + 9), heatgen_bytech_chart1)

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
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        heatgen_bytech_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_heatgen_3_rows):
            if not netz_heatgen_3['TECHNOLOGY'].iloc[i] in ['Gas CCS', 'Total']: 
                heatgen_bytech_chart2.add_series({
                    'name':       ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + i + 10, 0],
                    'categories': ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9, 2,\
                        (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9, netz_heatgen_3_cols - 1],
                    'values':     ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + i + 10, 2,\
                        (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + i + 10, netz_heatgen_3_cols - 1],
                    'fill':       {'color': netz_heatgen_3['TECHNOLOGY'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                if not netz_heatgen_2['TECHNOLOGY'].iloc[i] in ['Total']:
                    heatgen_bytech_chart2.add_series({
                        'name':       ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + i + 10, 0],
                        'categories': ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9, 2,\
                            (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9, netz_heatgen_3_cols - 1],
                        'values':     ['Heat generation', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + i + 10, 2,\
                            (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + i + 10, netz_heatgen_3_cols - 1],
                        'pattern':    {'fg_color': netz_heatgen_2['TECHNOLOGY'].map(colours_dict).loc[i],
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True},
                        'gap':        100
                    })

                else:
                    pass

        ref_worksheet27.insert_chart('J' + str(chart_height + ref_heatgen_2_rows + ref_heatgen_3_rows + 9), heatgen_bytech_chart2)
    
    else:
        pass

    #################################################################################

    # Create a use by fuel area chart
    if netz_heat_use_2_rows > 0:
        netz_heatuse_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_heatuse_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_heatuse_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_heatuse_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_heatuse_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_heatuse_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_heatuse_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_heat_use_2_rows):
            if not netz_heat_use_2['FUEL'].iloc[i] in ['Total']:
                netz_heatuse_chart1.add_series({
                    'name':       ['Heat only consumption', (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + i + 7, 0],
                    'categories': ['Heat only consumption', (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + 6, 2,\
                        (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + 6, netz_heat_use_2_cols - 1],
                    'values':     ['Heat only consumption', (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + i + 7, 2,\
                        (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + i + 7, netz_heat_use_2_cols - 1],
                    'fill':       {'color': netz_heat_use_2['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet28.insert_chart('B' + str(chart_height + ref_heat_use_2_rows + ref_heat_use_3_rows + 9), netz_heatuse_chart1)

    else:
        pass

    # Create a use column chart
    if netz_heat_use_3_rows > 0:
        netz_heatuse_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_heatuse_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_heatuse_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_heatuse_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_heatuse_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_heatuse_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_heatuse_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(netz_heat_use_3_rows):
            if not netz_heat_use_3['FUEL'].iloc[i] in ['Total']:
                netz_heatuse_chart2.add_series({
                    'name':       ['Heat only consumption', (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + netz_heat_use_2_rows + i + 10, 0],
                    'categories': ['Heat only consumption', (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + netz_heat_use_2_rows + 9, 2,\
                        (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + netz_heat_use_2_rows + 9, netz_heat_use_3_cols - 1],
                    'values':     ['Heat only consumption', (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + netz_heat_use_2_rows + i + 10, 2,\
                        (chart_height * 2) + ref_heat_use_2_rows + ref_heat_use_3_rows + netz_heat_use_2_rows + i + 10, netz_heat_use_3_cols - 1],
                    'fill':       {'color': netz_heat_use_3['FUEL'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        100
                })

            else:
                pass

        ref_worksheet28.insert_chart('J' + str(chart_height + ref_heat_use_2_rows + ref_heat_use_3_rows + 9), netz_heatuse_chart2)

    else:
        pass

    # Miscellaneous

    # Access the workbook and second sheet
    both_worksheet31 = writer.sheets['Modern renewables']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet31.set_column(2, ref_modren_4_cols + 1, None, space_format)
    both_worksheet31.set_row(chart_height, None, header_format)
    both_worksheet31.set_row(chart_height + ref_modren_4_rows + 3, None, header_format)
    both_worksheet31.set_row(chart_height + ref_modren_4_rows, None, percentage_format)
    both_worksheet31.set_row(chart_height + ref_modren_4_rows + netz_modren_4_rows + 3, None, percentage_format)
    both_worksheet31.write(0, 0, economy + ' modern renewables', cell_format1)

    # line chart
    if (ref_modren_4_rows > 0) & (netz_modren_4_rows > 0):
        modren_chart1 = workbook.add_chart({'type': 'line'})
        modren_chart1.set_size({
            'width': 500,            
            'height': 300
        })
            
        modren_chart1.set_chartarea({
            'border': {'none': True}
        })
            
        modren_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
                
        modren_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Proportion',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
                
        modren_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
                
        modren_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        # for component in ['Reference proportion', 'Carbon Neutrality proportion']:
        i = ref_modren_4[ref_modren_4['item_code_new'] == 'Reference'].index[0]
        modren_chart1.add_series({
            'name':       ['Modern renewables', chart_height + i + 1, 1],
            'categories': ['Modern renewables', chart_height, 2, chart_height, ref_modren_4_cols - 1],
            'values':     ['Modern renewables', chart_height + i + 1, 2, chart_height + i + 1, ref_modren_4_cols - 1],
            'line':       {'color': ref_modren_4['item_code_new'].map(colours_dict).loc[i],
                            'width': 1.5}
        })
        j = netz_modren_4[netz_modren_4['item_code_new'] == 'Carbon neutrality'].index[0]
        modren_chart1.add_series({
            'name':       ['Modern renewables', chart_height + ref_modren_4_rows + j + 4, 1],
            'categories': ['Modern renewables', chart_height + ref_modren_4_rows + 3, 2, chart_height + ref_modren_4_rows + 3, netz_modren_4_cols - 1],
            'values':     ['Modern renewables', chart_height + ref_modren_4_rows + j + 4, 2, chart_height + ref_modren_4_rows + j + 4, netz_modren_4_cols - 1],
            'line':       {'color': netz_modren_4['item_code_new'].map(colours_dict).loc[j],
                            'width': 1.5}
        })    
                
        both_worksheet31.insert_chart('B3', modren_chart1)

        # Stacked area electricity and heat
        modren_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'percent_stacked'})
        modren_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        modren_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        modren_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        modren_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Reference modern renewable electricity share',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        modren_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        modren_chart2.set_title({
            'none': True
        })

        # Configure the series of the chart from the dataframe data.    
        for option in ['Modern renewables', 'Non modern renewables']:
            i = ref_modren_4[(ref_modren_4['item_code_new'] == 'Electricity and heat') &
                             (ref_modren_4['fuel_code'] == option)].index[0]
            modren_chart2.add_series({
                'name':       ['Modern renewables', chart_height + i + 1, 0],
                'categories': ['Modern renewables', chart_height, 2, chart_height, ref_modren_4_cols - 1],
                'values':     ['Modern renewables', chart_height + i + 1, 2, chart_height + i + 1, ref_modren_4_cols - 1],
                'fill':       {'color': ref_modren_4['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })
        
        both_worksheet31.insert_chart('J3', modren_chart2)

        # Stacked area electricity and heat
        modren_chart3 = workbook.add_chart({'type': 'area', 'subtype': 'percent_stacked'})
        modren_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        modren_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        modren_chart3.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        modren_chart3.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Carbon Neutrality modern renewable electricity share',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        modren_chart3.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        modren_chart3.set_title({
            'none': True
        })

        # Configure the series of the chart from the dataframe data.    
        for option in ['Modern renewables', 'Non modern renewables']:
            i = netz_modren_4[(netz_modren_4['item_code_new'] == 'Electricity and heat') &
                             (netz_modren_4['fuel_code'] == option)].index[0]
            modren_chart3.add_series({
                'name':       ['Modern renewables', chart_height + ref_modren_4_rows + i + 4, 0],
                'categories': ['Modern renewables', chart_height + ref_modren_4_rows + 3, 2, chart_height + ref_modren_4_rows + 3, netz_modren_4_cols - 1],
                'values':     ['Modern renewables', chart_height + ref_modren_4_rows + i + 4, 2, chart_height + ref_modren_4_rows + i + 4, netz_modren_4_cols - 1],
                'fill':       {'color': netz_modren_4['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })
        
        both_worksheet31.insert_chart('R3', modren_chart3)
    
    else:
        pass

    ##############################################################
    # Energy intensity chart

    # Access the workbook and second sheet
    both_worksheet37 = writer.sheets['Energy intensity']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet37.set_column(2, ref_enint_3_cols + 1, None, space_format)
    both_worksheet37.set_row(chart_height, None, header_format)
    both_worksheet37.set_row(chart_height + ref_enint_3_rows + 3, None, header_format)
    both_worksheet37.set_row(chart_height + ref_enint_3_rows + netz_enint_3_rows + 6, None, header_format)
    both_worksheet37.set_row(chart_height + ref_enint_3_rows + netz_enint_3_rows + ref_enint_sup3_rows + 9, None, header_format)
    both_worksheet37.write(0, 0, economy + ' energy intensity', cell_format1)

    # line chart
    if (ref_enint_3_rows > 0) & (netz_enint_3_rows > 0):
        enint_chart1 = workbook.add_chart({'type': 'line'})
        enint_chart1.set_size({
            'width': 500,            
            'height': 300
        })
            
        enint_chart1.set_chartarea({
            'border': {'none': True}
        })
            
        enint_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
                
        enint_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'TFEC energy intensity (2005 = 100)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
                
        enint_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
                
        enint_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        i = ref_enint_3[ref_enint_3['Series'] == 'Reference'].index[0]
        enint_chart1.add_series({
            'name':       ['Energy intensity', chart_height + i + 1, 1],
            'categories': ['Energy intensity', chart_height, 2, chart_height, ref_enint_3_cols - 1],
            'values':     ['Energy intensity', chart_height + i + 1, 2, chart_height + i + 1, ref_enint_3_cols - 1],
            'line':       {'color': ref_enint_3['Series'].map(colours_dict).loc[i],
                            'width': 1.5}
        })
        j = netz_enint_3[netz_enint_3['Series'] == 'Carbon neutrality'].index[0]
        enint_chart1.add_series({
            'name':       ['Energy intensity', chart_height + ref_enint_3_rows + j + 4, 1],
            'categories': ['Energy intensity', chart_height + ref_enint_3_rows + 3, 2, chart_height + ref_enint_3_rows + 3, netz_enint_3_cols - 1],
            'values':     ['Energy intensity', chart_height + ref_enint_3_rows + j + 4, 2, chart_height + ref_enint_3_rows + j + 4, netz_enint_3_cols - 1],
            'line':       {'color': netz_enint_3['Series'].map(colours_dict).loc[j],
                            'width': 1.5}
        })
        if economy == 'APEC':
            k = ref_enint_3[ref_enint_3['Series'] == 'Target'].index[0]
            enint_chart1.add_series({
                'name':       ['Energy intensity', chart_height + k + 1, 1],
                'categories': ['Energy intensity', chart_height, 2, chart_height, ref_enint_3_cols - 1],
                'values':     ['Energy intensity', chart_height + k + 1, 2, chart_height + k + 1, ref_enint_3_cols - 1],
                'line':       {'color': ref_enint_3['Series'].map(colours_dict).loc[k],
                                'width': 1.5}
            })    
                
        both_worksheet37.insert_chart('B3', enint_chart1)

    else:
        pass

    # line chart
    if (ref_enint_sup3_rows > 0) & (netz_enint_sup3_rows > 0):
        enint_chart2 = workbook.add_chart({'type': 'line'})
        enint_chart2.set_size({
            'width': 500,            
            'height': 300
        })
            
        enint_chart2.set_chartarea({
            'border': {'none': True}
        })
            
        enint_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
                
        enint_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'TPES energy intensity',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0.0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
                
        enint_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
                
        enint_chart2.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        i = ref_enint_sup3[ref_enint_sup3['Series'] == 'Reference'].index[0]
        enint_chart2.add_series({
            'name':       ['Energy intensity', chart_height + ref_enint_3_rows + netz_enint_3_rows + i + 7, 1],
            'categories': ['Energy intensity', chart_height + ref_enint_3_rows + netz_enint_3_rows + 6, 2,\
                chart_height + ref_enint_3_rows + netz_enint_3_rows + 6, ref_enint_sup3_cols - 1],
            'values':     ['Energy intensity', chart_height + ref_enint_3_rows + netz_enint_3_rows + i + 7, 2,\
                chart_height + ref_enint_3_rows + netz_enint_3_rows + i + 7, ref_enint_sup3_cols - 1],
            'line':       {'color': ref_enint_sup3['Series'].map(colours_dict).loc[i],
                            'width': 1.5}

            })

        j = netz_enint_sup3[netz_enint_sup3['Series'] == 'Carbon neutrality'].index[0]
        enint_chart2.add_series({
            'name':       ['Energy intensity', chart_height + ref_enint_3_rows + netz_enint_3_rows + ref_enint_sup3_rows + j + 10, 1],
            'categories': ['Energy intensity', chart_height + ref_enint_3_rows + netz_enint_3_rows + ref_enint_sup3_rows + 9, 2,\
                chart_height + ref_enint_3_rows + netz_enint_3_rows + ref_enint_sup3_rows + 9, netz_enint_sup3_cols - 1],
            'values':     ['Energy intensity', chart_height + ref_enint_3_rows + netz_enint_3_rows + ref_enint_sup3_rows + j + 10, 2,\
                chart_height + ref_enint_3_rows + netz_enint_3_rows + ref_enint_sup3_rows + j + 10, netz_enint_sup3_cols - 1],
            'line':       {'color': netz_enint_sup3['Series'].map(colours_dict).loc[j],
                            'width': 1.5}

            })
                        
        both_worksheet37.insert_chart('J3', enint_chart2)

    else:
        pass

    ################################################
    # Macro charts

    # Access the workbook and second sheet
    both_worksheet32 = writer.sheets['Macro']
    
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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
                
        GDP_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'GDP (billions 2018 USD PPP 2018)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
                
        GDP_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9},
            'none': True
        })
                
        GDP_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        # for component in ['Reference proportion', 'Net-zero proportion']:
        i = macro_1[macro_1['Series'] == 'GDP 2018 USD PPP'].index[0]
        GDP_chart1.add_series({
            'name':       ['Macro', chart_height + i + 1, 1],
            'categories': ['Macro', chart_height, 2, chart_height, macro_1_cols - 1],
            'values':     ['Macro', chart_height + i + 1, 2, chart_height + i + 1, macro_1_cols - 1],
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
                # 'name': 'Year',
                'label_position': 'low',
                'major_tick_mark': 'none',
                'minor_tick_mark': 'none',
                'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
                #'position_axis': 'on_tick',
                'interval_unit': 10,
                'line': {'color': '#bebebe'}
            })
                    
            GDP_chart2.set_y_axis({
                'major_tick_mark': 'none', 
                'minor_tick_mark': 'none',
                'name': 'GDP growth',
                'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
                'num_format': '0%',
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': '#bebebe'}
                },
                'line': {'color': '#bebebe'}
            })
                    
            GDP_chart2.set_legend({
                'font': {'name': 'Segoe UI', 'size': 9},
                'none': True
            })
                    
            GDP_chart2.set_title({
                'none': True
            })
                
            # Configure the series of the chart from the dataframe data.
            i = macro_1[macro_1['Series'] == 'GDP growth'].index[0]
            GDP_chart2.add_series({
                'name':       ['Macro', chart_height + i + 1, 1],
                'categories': ['Macro', chart_height, 2, chart_height, macro_1_cols - 1],
                'values':     ['Macro', chart_height + i + 1, 2, chart_height + i + 1, macro_1_cols - 1],
                'fill':       {'color': macro_1['Series'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
                
        pop_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'Population (millions)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
                
        pop_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9},
            'none': True
        })
                
        pop_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        # for component in ['Reference proportion', 'Net-zero proportion']:
        i = macro_1[macro_1['Series'] == 'Population'].index[0]
        pop_chart1.add_series({
            'name':       ['Macro', chart_height + i + 1, 1],
            'categories': ['Macro', chart_height, 2, chart_height, macro_1_cols - 1],
            'values':     ['Macro', chart_height + i + 1, 2, chart_height + i + 1, macro_1_cols - 1],
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
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
                
        GDPpc_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            'name': 'GDP per capita (2018 USD PPP)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
                
        GDPpc_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9},
            'none': True
        })
                
        GDPpc_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        # for component in ['Reference proportion', 'Net-zero proportion']:
        i = macro_1[macro_1['Series'] == 'GDP per capita'].index[0]
        GDPpc_chart1.add_series({
            'name':       ['Macro', chart_height + i + 1, 1],
            'categories': ['Macro', chart_height, 2, chart_height, macro_1_cols - 1],
            'values':     ['Macro', chart_height + i + 1, 2, chart_height + i + 1, macro_1_cols - 1],
            'line':       {'color': macro_1['Series'].map(colours_dict).loc[i],
                            'width': 1.5}
        })

        both_worksheet32.insert_chart('Z3', GDPpc_chart1)

    else:
        pass

    ################################################
    # Heavy industry

    # Access the workbook and second sheet
    both_worksheet33 = writer.sheets['Heavy industry']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet33.set_column(2, macro_1_cols + 1, None, space_format)
    both_worksheet33.set_row(chart_height, None, header_format)
    both_worksheet33.set_row(chart_height + ref_steel_3_rows + 3, None, header_format)
    both_worksheet33.set_row(chart_height + ref_steel_3_rows + ref_chem_3_rows + 6, None, header_format)
    both_worksheet33.set_row((2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + 9, None, header_format)
    both_worksheet33.set_row((2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + 12, None, header_format)
    both_worksheet33.set_row((2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + netz_chem_3_rows + 15, None, header_format)
    both_worksheet33.write(0, 0, economy + ' heavy industry fuel use Reference', cell_format1)
    both_worksheet33.write(chart_height + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + 9, 0,\
        economy + ' heavy industry fuel use Carbon Neutrality', cell_format1)
    both_worksheet33.write(1, 0, 'Units: Petajoules', cell_format2)

    # Steel stacked chart
    if ref_steel_3_rows > 0:
        ref_steel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_steel_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_steel_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_steel_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_steel_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_steel_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_steel_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_steel_3_rows):
            ref_steel_chart1.add_series({
                'name':       ['Heavy industry', chart_height + i + 1, 1],
                'categories': ['Heavy industry', chart_height, 2, chart_height, ref_steel_3_cols - 1],
                'values':     ['Heavy industry', chart_height + i + 1, 2, chart_height + i + 1, ref_steel_3_cols - 1],
                'fill':       {'color': ref_steel_3['tech_mix'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet33.insert_chart('B3', ref_steel_chart1)

    else: 
        pass

    # Chemicals stacked chart
    if ref_chem_3_rows > 0:
        ref_chem_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_chem_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_chem_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_chem_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_chem_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_chem_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_chem_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_chem_3_rows):
            ref_chem_chart1.add_series({
                'name':       ['Heavy industry', chart_height + ref_steel_3_rows + i + 4, 1],
                'categories': ['Heavy industry', chart_height + ref_steel_3_rows + 3, 2, chart_height + ref_steel_3_rows + 3, ref_chem_3_cols - 1],
                'values':     ['Heavy industry', chart_height + ref_steel_3_rows + i + 4, 2, chart_height + ref_steel_3_rows + i + 4, ref_chem_3_cols - 1],
                'fill':       {'color': ref_chem_3['tech_mix'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet33.insert_chart('J3', ref_chem_chart1)

    else: 
        pass

    # Cement stacked chart
    if ref_cement_3_rows > 0:
        ref_cement_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_cement_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_cement_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_cement_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_cement_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_cement_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_cement_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_cement_3_rows):
            ref_cement_chart1.add_series({
                'name':       ['Heavy industry', chart_height + ref_steel_3_rows + ref_chem_3_rows + i + 7, 1],
                'categories': ['Heavy industry', chart_height + ref_steel_3_rows + ref_chem_3_rows + 6, 2,\
                    chart_height + ref_steel_3_rows + ref_chem_3_rows + 6, ref_cement_3_cols - 1],
                'values':     ['Heavy industry', chart_height + ref_steel_3_rows + ref_chem_3_rows + i + 7, 2,\
                    chart_height + ref_steel_3_rows + ref_chem_3_rows + i + 7, ref_cement_3_cols - 1],
                'fill':       {'color': ref_cement_3['tech_mix'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet33.insert_chart('R3', ref_cement_chart1)

    else: 
        pass

    # NZS Steel stacked chart
    if netz_steel_3_rows > 0:
        netz_steel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_steel_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_steel_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_steel_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_steel_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_steel_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_steel_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_steel_3_rows):
            netz_steel_chart1.add_series({
                'name':       ['Heavy industry', (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + i + 10, 1],
                'categories': ['Heavy industry', (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + 9, 2,\
                    (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + 9, netz_steel_3_cols - 1],
                'values':     ['Heavy industry', (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + i + 10, 2,\
                    (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + i + 10, netz_steel_3_cols - 1],
                'fill':       {'color': netz_steel_3['tech_mix'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet33.insert_chart('B' + str(chart_height + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + 12), netz_steel_chart1)

    else: 
        pass

    # NZS Chemicals stacked chart
    if netz_chem_3_rows > 0:
        netz_chem_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_chem_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_chem_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_chem_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_chem_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_chem_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_chem_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_chem_3_rows):
            netz_chem_chart1.add_series({
                'name':       ['Heavy industry', (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + i + 13, 1],
                'categories': ['Heavy industry', (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + 12, 2,\
                    (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + 12, netz_chem_3_cols - 1],
                'values':     ['Heavy industry', (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + i + 13, 2,\
                    (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + i + 13, netz_chem_3_cols - 1],
                'fill':       {'color': netz_chem_3['tech_mix'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet33.insert_chart('J' + str(chart_height + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + 12), netz_chem_chart1)

    else: 
        pass

    # NZS Cement stacked chart
    if netz_cement_3_rows > 0:
        netz_cement_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_cement_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_cement_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_cement_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_cement_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_cement_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_cement_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_cement_3_rows):
            netz_cement_chart1.add_series({
                'name':       ['Heavy industry', (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + netz_chem_3_rows + i + 16, 1],
                'categories': ['Heavy industry', (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + netz_chem_3_rows + 15, 2,\
                    (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + netz_chem_3_rows + 15, netz_cement_3_cols - 1],
                'values':     ['Heavy industry', (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + netz_chem_3_rows + i + 16, 2,\
                    (2 * chart_height) + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + netz_steel_3_rows + netz_chem_3_rows + i + 16, netz_cement_3_cols - 1],
                'fill':       {'color': netz_cement_3['tech_mix'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet33.insert_chart('R' + str(chart_height + ref_steel_3_rows + ref_chem_3_rows + ref_cement_3_rows + 12), netz_cement_chart1)

    else: 
        pass

    ##################################################################################

    # Emissions charts

    # Access the workbook and first sheet with data from df1
    both_worksheet34 = writer.sheets['CO2 by fuel']
        
    # Apply comma format and header format to relevant data rows
    both_worksheet34.set_column(1, ref_emiss_fuel_1_cols + 1, None, space_format)
    both_worksheet34.set_row(chart_height, None, header_format)
    both_worksheet34.set_row(chart_height + ref_emiss_fuel_1_rows + 3, None, header_format)
    both_worksheet34.set_row((2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, None, header_format)
    both_worksheet34.set_row((2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + 9, None, header_format)
    both_worksheet34.write(0, 0, economy + ' emissions by fuel Reference scenario', cell_format1)
    both_worksheet34.write(chart_height + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, 0, economy + ' emissions by fuel Carbon Neutrality scenario', cell_format1)
    both_worksheet34.write(1, 0, 'Units: Million tonnes of CO2', cell_format2)

    ################################################################### CHARTS ###################################################################

    # Create a FED area chart
    ref_em_fuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_em_fuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_fuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_fuel_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        #'name': 'Million tonnes CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_em_fuel_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_em_fuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_emiss_fuel_1_rows):
        if not ref_emiss_fuel_1['fuel_code'].iloc[i] in ['Total']:
            ref_em_fuel_chart1.add_series({
                'name':       ['CO2 by fuel', chart_height + i + 1, 0],
                'categories': ['CO2 by fuel', chart_height, 2, chart_height, ref_emiss_fuel_1_cols - 1],
                'values':     ['CO2 by fuel', chart_height + i + 1, 2, chart_height + i + 1, ref_emiss_fuel_1_cols - 1],
                'fill':       {'color': ref_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
    both_worksheet34.insert_chart('B3', ref_em_fuel_chart1)

    ###################### Create another EMISSIONS chart showing proportional share #################################

    # Create a another chart
    ref_em_fuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    ref_em_fuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_fuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_fuel_chart2.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        #'name': 'CO2 proportion',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_em_fuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in ref_emiss_fuel_2['fuel_code'].unique():
        i = ref_emiss_fuel_2[ref_emiss_fuel_2['fuel_code'] == component].index[0]
        if not ref_emiss_fuel_2['fuel_code'].iloc[i] in ['Total']:
            ref_em_fuel_chart2.add_series({
                'name':       ['CO2 by fuel', chart_height + ref_emiss_fuel_1_rows + i + 4, 0],
                'categories': ['CO2 by fuel', chart_height + ref_emiss_fuel_1_rows + 3, 2, chart_height + ref_emiss_fuel_1_rows + 3, ref_emiss_fuel_2_cols - 1],
                'values':     ['CO2 by fuel', chart_height + ref_emiss_fuel_1_rows + i + 4, 2, chart_height + ref_emiss_fuel_1_rows + i + 4, ref_emiss_fuel_2_cols - 1],
                'fill':       {'color': ref_emiss_fuel_2['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })

        else:
            pass
    
    both_worksheet34.insert_chart('J3', ref_em_fuel_chart2)

    # Create a Emissions line chart with higher level aggregation
    ref_em_fuel_chart3 = workbook.add_chart({'type': 'line'})
    ref_em_fuel_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_fuel_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_fuel_chart3.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        #'name': 'Million tonnes CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_em_fuel_chart3.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_em_fuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_emiss_fuel_1_rows):
        if not ref_emiss_fuel_1['fuel_code'].iloc[i] in ['Total']:
            ref_em_fuel_chart3.add_series({
                'name':       ['CO2 by fuel', chart_height + i + 1, 0],
                'categories': ['CO2 by fuel', chart_height, 2, chart_height, ref_emiss_fuel_1_cols - 1],
                'values':     ['CO2 by fuel', chart_height + i + 1, 2, chart_height + i + 1, ref_emiss_fuel_1_cols - 1],
                'line':       {'color': ref_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i], 
                            'width': 1}
            })

        else:
            ref_em_fuel_chart3.add_series({
                'name':       ['CO2 by fuel', chart_height + i + 1, 0],
                'categories': ['CO2 by fuel', chart_height, 2, chart_height, ref_emiss_fuel_1_cols - 1],
                'values':     ['CO2 by fuel', chart_height + i + 1, 2, chart_height + i + 1, ref_emiss_fuel_1_cols - 1],
                'line':       {'color': ref_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i], 
                            'width': 1.5}
            })    
        
    both_worksheet34.insert_chart('R3', ref_em_fuel_chart3)

    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    both_worksheet35 = writer.sheets['CO2 by sector']
        
    # Apply comma format and header format to relevant data rows
    both_worksheet35.set_column(1, ref_emiss_2_cols + 1, None, space_format)
    both_worksheet35.set_row(chart_height, None, header_format)
    both_worksheet35.set_row(chart_height + ref_emiss_sector_1_rows + 3, None, header_format)
    both_worksheet35.set_row((2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, None, header_format)
    both_worksheet35.set_row((2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + 9, None, header_format)
    both_worksheet35.write(0, 0, economy + ' emissions by demand sector Reference scenario', cell_format1)
    both_worksheet35.write(chart_height + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, 0, economy + ' emissions by demand sector Carbon Neutrality scenario', cell_format1)
    both_worksheet35.write(1, 0, 'Units: Million tonnes of CO2', cell_format2)
    
    # Create an EMISSIONS sector line chart

    ref_em_sector_chart1 = workbook.add_chart({'type': 'line'})
    ref_em_sector_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_sector_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_sector_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        #'name': 'Million tonnes CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_em_sector_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_em_sector_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_emiss_sector_1_rows):
        if not ref_emiss_sector_1['item_code_new'].iloc[i] in ['Total']:
            ref_em_sector_chart1.add_series({
                'name':       ['CO2 by sector', chart_height + i + 1, 1],
                'categories': ['CO2 by sector', chart_height, 2, chart_height, ref_emiss_sector_1_cols - 1],
                'values':     ['CO2 by sector', chart_height + i + 1, 2, chart_height + i + 1, ref_emiss_sector_1_cols - 1],
                'line':       {'color': ref_emiss_sector_1['item_code_new'].map(colours_dict).loc[i], 
                            'width': 1}
            })

        else:
            pass    
        
    both_worksheet35.insert_chart('R3', ref_em_sector_chart1)

    # Create a EMISSIONS sector area chart

    ref_em_sector_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_em_sector_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_sector_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_sector_chart2.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        #'name': 'Million tonnes CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    ref_em_sector_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_em_sector_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_emiss_sector_1_rows):
        if not ref_emiss_sector_1['item_code_new'].iloc[i] in ['Total']:
            ref_em_sector_chart2.add_series({
                'name':       ['CO2 by sector', chart_height + i + 1, 1],
                'categories': ['CO2 by sector', chart_height, 2, chart_height, ref_emiss_sector_1_cols - 1],
                'values':     ['CO2 by sector', chart_height + i + 1, 2, chart_height + i + 1, ref_emiss_sector_1_cols - 1],
                'fill':       {'color': ref_emiss_sector_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
    both_worksheet35.insert_chart('B3', ref_em_sector_chart2)

    ###################### Create another FED chart showing proportional share #################################

    # Create a FED chart
    ref_em_sector_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    ref_em_sector_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_sector_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_sector_chart3.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        #'name': 'Proportion of CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart3.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_em_sector_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in ref_emiss_sector_2['item_code_new'].unique():
        i = ref_emiss_sector_2[ref_emiss_sector_2['item_code_new'] == component].index[0]
        if not ref_emiss_sector_2['item_code_new'].iloc[i] in ['Total']:
            ref_em_sector_chart3.add_series({
                'name':       ['CO2 by sector', chart_height + ref_emiss_sector_1_rows + i + 4, 1],
                'categories': ['CO2 by sector', chart_height + ref_emiss_sector_1_rows + 3, 2, chart_height + ref_emiss_sector_1_rows + 3, ref_emiss_sector_2_cols - 1],
                'values':     ['CO2 by sector', chart_height + ref_emiss_sector_1_rows + i + 4, 2, chart_height + ref_emiss_sector_1_rows + i + 4, ref_emiss_sector_2_cols - 1],
                'fill':       {'color': ref_emiss_sector_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })

        else:
            pass
    
    both_worksheet35.insert_chart('J3', ref_em_sector_chart3)

    #############################################################################################################################
    # NET ZERO CHARTS EMISSIONS
    ################################################################### CHARTS ##################################################

    # Create a FED area chart
    netz_em_fuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_em_fuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_fuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_fuel_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        #'name': 'Million tonnes CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_em_fuel_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_em_fuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_emiss_fuel_1_rows):
        if not netz_emiss_fuel_1['fuel_code'].iloc[i] in ['Total']:
            netz_em_fuel_chart1.add_series({
                'name':       ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 0],
                'categories': ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, 2,\
                    (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, netz_emiss_fuel_1_cols - 1],
                'values':     ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, netz_emiss_fuel_1_cols - 1],
                'fill':       {'color': netz_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
    both_worksheet34.insert_chart('B' + str(chart_height + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 9), netz_em_fuel_chart1)

    ###################### Create another EMISSIONS chart showing proportional share #################################

    # Create a another chart
    netz_em_fuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    netz_em_fuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_fuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_fuel_chart2.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        #'name': 'CO2 proportion',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_em_fuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in netz_emiss_fuel_2['fuel_code'].unique():
        i = netz_emiss_fuel_2[netz_emiss_fuel_2['fuel_code'] == component].index[0]
        if not netz_emiss_fuel_2['fuel_code'].iloc[i] in ['Total']:
            netz_em_fuel_chart2.add_series({
                'name':       ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + i + 10, 0],
                'categories': ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + 9, 2,\
                    (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + 9, netz_emiss_fuel_2_cols - 1],
                'values':     ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + i + 10, netz_emiss_fuel_2_cols - 1],
                'fill':       {'color': netz_emiss_fuel_2['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })

        else:
            pass
    
    both_worksheet34.insert_chart('J' + str(chart_height + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 9), netz_em_fuel_chart2)

    # Create a Emissions line chart with higher level aggregation
    netz_em_fuel_chart3 = workbook.add_chart({'type': 'line'})
    netz_em_fuel_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_fuel_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_fuel_chart3.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        #'name': 'Million tonnes CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_em_fuel_chart3.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_em_fuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_emiss_fuel_1_rows):
        if not netz_emiss_fuel_1['fuel_code'].iloc[i] in ['Total']:
            netz_em_fuel_chart3.add_series({
                'name':       ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 0],
                'categories': ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, 2,\
                    (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, netz_emiss_fuel_1_cols - 1],
                'values':     ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, netz_emiss_fuel_1_cols - 1],
                'line':       {'color': netz_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i], 
                            'width': 1}
            })

        else:
            netz_em_fuel_chart3.add_series({
                'name':       ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 0],
                'categories': ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, 2,\
                    (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, netz_emiss_fuel_1_cols - 1],
                'values':     ['CO2 by fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, netz_emiss_fuel_1_cols - 1],
                'line':       {'color': netz_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i], 
                            'width': 1.5}
            })
        
    both_worksheet34.insert_chart('R' + str(chart_height + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 9), netz_em_fuel_chart3)

    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Create an EMISSIONS sector line chart

    netz_em_sector_chart1 = workbook.add_chart({'type': 'line'})
    netz_em_sector_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_sector_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_sector_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        #'name': 'Million tonnes CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_em_sector_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_em_sector_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_emiss_sector_1_rows):
        if not netz_emiss_sector_1['item_code_new'].iloc[i] in ['Total']:
            netz_em_sector_chart1.add_series({
                'name':       ['CO2 by sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, 1],
                'categories': ['CO2 by sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, 2,\
                    (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, netz_emiss_sector_1_cols - 1],
                'values':     ['CO2 by sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, netz_emiss_sector_1_cols - 1],
                'line':       {'color': netz_emiss_sector_1['item_code_new'].map(colours_dict).loc[i], 
                            'width': 1}
            })

        else:
            pass    
        
    both_worksheet35.insert_chart('R' + str(chart_height + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 9), netz_em_sector_chart1)

    # Create a EMISSIONS sector area chart

    netz_em_sector_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_em_sector_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_sector_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_sector_chart2.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        #'name': 'Million tonnes CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    netz_em_sector_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_em_sector_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_emiss_sector_1_rows):
        if not netz_emiss_sector_1['item_code_new'].iloc[i] in ['Total']:
            netz_em_sector_chart2.add_series({
                'name':       ['CO2 by sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, 1],
                'categories': ['CO2 by sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, 2,\
                    (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, netz_emiss_sector_1_cols - 1],
                'values':     ['CO2 by sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, netz_emiss_sector_1_cols - 1],
                'fill':       {'color': netz_emiss_sector_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        else:
            pass    
        
    both_worksheet35.insert_chart('B' + str(chart_height + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 9), netz_em_sector_chart2)

    ###################### Create another FED chart showing proportional share #################################

    # Create a FED chart
    netz_em_sector_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    netz_em_sector_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_sector_chart3.set_chartarea({
        'border': {'none': True} 
    })
    
    netz_em_sector_chart3.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        #'name': 'Proportion of CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart3.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_em_sector_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in netz_emiss_sector_2['item_code_new'].unique():
        i = netz_emiss_sector_2[netz_emiss_sector_2['item_code_new'] == component].index[0]
        if not netz_emiss_sector_2['item_code_new'].iloc[i] in ['Total']:
            netz_em_sector_chart3.add_series({
                'name':       ['CO2 by sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + i + 10, 1],
                'categories': ['CO2 by sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + 9, 2,\
                    (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + 9, netz_emiss_sector_2_cols - 1],
                'values':     ['CO2 by sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + i + 10, netz_emiss_sector_2_cols - 1],
                'fill':       {'color': netz_emiss_sector_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })

        else:
            pass
    
    both_worksheet35.insert_chart('J' + str(chart_height + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 9), netz_em_sector_chart3)

    ################################################
    # Transport road

    # Access the workbook and second sheet
    both_worksheet36 = writer.sheets['Road transport']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet36.set_column(2, ref_roadmodal_3_cols + 1, None, space_format)
    both_worksheet36.set_row(chart_height, None, header_format)
    both_worksheet36.set_row(chart_height + ref_roadmodal_3_rows + 3, None, header_format)
    both_worksheet36.set_row((2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + 6, None, header_format)
    both_worksheet36.set_row((2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + netz_roadmodal_3_rows + 9, None, header_format)
    both_worksheet36.write(0, 0, economy + ' road transport Reference', cell_format1)
    both_worksheet36.write(chart_height + ref_roadmodal_3_rows + ref_roadfuel_3_rows + 6, 0,\
        economy + ' road transport Carbon Neutrality', cell_format1)
    both_worksheet36.write(1, 0, 'Units: Petajoules', cell_format2)

    # Road modal area stacked chart
    if ref_roadmodal_3_rows > 0:
        ref_roadmodal_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_roadmodal_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_roadmodal_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_roadmodal_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_roadmodal_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_roadmodal_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_roadmodal_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_roadmodal_3_rows):
            ref_roadmodal_chart1.add_series({
                'name':       ['Road transport', chart_height + i + 1, 1],
                'categories': ['Road transport', chart_height, 2, chart_height, ref_roadmodal_3_cols - 1],
                'values':     ['Road transport', chart_height + i + 1, 2, chart_height + i + 1, ref_roadmodal_3_cols - 1],
                'fill':       {'color': ref_roadmodal_3['modality'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet36.insert_chart('B3', ref_roadmodal_chart1)

    else: 
        pass

    # Road modal area stacked chart
    if ref_roadfuel_3_rows > 0:
        ref_roadfuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_roadfuel_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_roadfuel_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_roadfuel_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_roadfuel_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_roadfuel_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_roadfuel_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_roadfuel_3_rows):
            ref_roadfuel_chart1.add_series({
                'name':       ['Road transport', chart_height + ref_roadmodal_3_rows + i + 4, 1],
                'categories': ['Road transport', chart_height + ref_roadmodal_3_rows + 3, 2,\
                    chart_height + ref_roadmodal_3_rows + 3, ref_roadfuel_3_cols - 1],
                'values':     ['Road transport', chart_height + ref_roadmodal_3_rows + i + 4, 2,\
                    chart_height + ref_roadmodal_3_rows + i + 4, ref_roadfuel_3_cols - 1],
                'fill':       {'color': ref_roadfuel_3['modality'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet36.insert_chart('J3', ref_roadfuel_chart1)

    else: 
        pass
    
    # CARBON NEUTRALITY

    # Road modal area stacked chart
    if netz_roadmodal_3_rows > 0:
        netz_roadmodal_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_roadmodal_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_roadmodal_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_roadmodal_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_roadmodal_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_roadmodal_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_roadmodal_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_roadmodal_3_rows):
            netz_roadmodal_chart1.add_series({
                'name':       ['Road transport', (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + i + 7, 1],
                'categories': ['Road transport', (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + 6, 2,\
                    (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + 6, netz_roadmodal_3_cols - 1],
                'values':     ['Road transport', (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + i + 7, 2,\
                    (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + i + 7, netz_roadmodal_3_cols - 1],
                'fill':       {'color': netz_roadmodal_3['modality'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet36.insert_chart('B' + str(chart_height + ref_roadmodal_3_rows + ref_roadfuel_3_rows + 9), netz_roadmodal_chart1)

    else: 
        pass

    # Road modal area stacked chart
    if netz_roadfuel_3_rows > 0:
        netz_roadfuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_roadfuel_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_roadfuel_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_roadfuel_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_roadfuel_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_roadfuel_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_roadfuel_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_roadfuel_3_rows):
            netz_roadfuel_chart1.add_series({
                'name':       ['Road transport', (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + netz_roadmodal_3_rows + i + 10, 1],
                'categories': ['Road transport', (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + netz_roadmodal_3_rows + 9, 2,\
                    (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + netz_roadmodal_3_rows + 9, netz_roadfuel_3_cols - 1],
                'values':     ['Road transport', (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + netz_roadmodal_3_rows + i + 10, 2,\
                    (2 * chart_height) + ref_roadmodal_3_rows + ref_roadfuel_3_rows + netz_roadmodal_3_rows + i + 10, netz_roadfuel_3_cols - 1],
                'fill':       {'color': netz_roadfuel_3['modality'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        both_worksheet36.insert_chart('J' + str(chart_height + ref_roadmodal_3_rows + ref_roadfuel_3_rows + 9), netz_roadfuel_chart1)

    else: 
        pass

    ##############################################################

    # TPES by fuel

    # access the sheet for production created above
    ref_worksheet15 = writer.sheets['Other fuels REF']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet15.set_column(2, ref_coal_1_cols + 1, None, space_format)
    ref_worksheet15.set_row(chart_height, None, header_format)
    ref_worksheet15.set_row(chart_height + ref_nuke_1_rows + 3, None, header_format)
    ref_worksheet15.set_row(chart_height + ref_nuke_1_rows + ref_biomass_1_rows + 6, None, header_format)
    ref_worksheet15.write(0, 0, economy + ' TPES nuclear, biomass and biofuels Reference', cell_format1)
    ref_worksheet15.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a TPES nuclear  chart
    if ref_nuke_1_rows > 0:
        ref_tpes_nuke_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_nuke_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_nuke_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_nuke_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_nuke_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Nuclear (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_nuke_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_nuke_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in ref_nuke_1['item_code_new'].unique():
            i = ref_nuke_1[ref_nuke_1['item_code_new'] == component].index[0]
            ref_tpes_nuke_chart1.add_series({
                'name':       ['Other fuels REF', chart_height + i + 1, 1],
                'categories': ['Other fuels REF', chart_height, 2, chart_height, ref_nuke_1_cols - 1],
                'values':     ['Other fuels REF', chart_height + i + 1, 2, chart_height + i + 1, ref_nuke_1_cols - 1],
                'fill':       {'color': ref_nuke_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet15.insert_chart('B3', ref_tpes_nuke_chart1)

    else:
        pass

    # Create a TPES biomass chart
    ref_tpes_biomass_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    ref_tpes_biomass_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_tpes_biomass_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_tpes_biomass_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    ref_tpes_biomass_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Biomass (PJ)',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_tpes_biomass_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_tpes_biomass_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in fuel_final_nobunk[:-1]:
        i = ref_biomass_1[ref_biomass_1['item_code_new'] == component].index[0]
        ref_tpes_biomass_chart1.add_series({
            'name':       ['Other fuels REF', chart_height + ref_nuke_1_rows + i + 4, 1],
            'categories': ['Other fuels REF', chart_height + ref_nuke_1_rows + 3, 2,\
                chart_height + ref_nuke_1_rows + 3, ref_biomass_1_cols - 1],
            'values':     ['Other fuels REF', chart_height + ref_nuke_1_rows + i + 4, 2,\
                chart_height + ref_nuke_1_rows + i + 4, ref_biomass_1_cols - 1],
            'fill':       {'color': ref_biomass_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True},
            'gap':        100
        })
    
    ref_worksheet15.insert_chart('J3', ref_tpes_biomass_chart1)

    # Create a TPES biofuel chart
    ref_tpes_biofuel_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    ref_tpes_biofuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_tpes_biofuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_tpes_biofuel_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    ref_tpes_biofuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Biofuels (PJ)',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_tpes_biofuel_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    ref_tpes_biofuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in fuel_final_bunk[:-1]:
        i = ref_biofuel_2[ref_biofuel_2['item_code_new'] == component].index[0]
        ref_tpes_biofuel_chart1.add_series({
            'name':       ['Other fuels REF', chart_height + ref_nuke_1_rows + ref_biomass_1_rows + i + 7, 1],
            'categories': ['Other fuels REF', chart_height + ref_nuke_1_rows + ref_biomass_1_rows + 6, 2,\
                chart_height + ref_nuke_1_rows + ref_biomass_1_rows + 6, ref_biofuel_2_cols - 1],
            'values':     ['Other fuels REF', chart_height + ref_nuke_1_rows + ref_biomass_1_rows + i + 7, 2,\
                chart_height + ref_nuke_1_rows + ref_biomass_1_rows + i + 7, ref_biofuel_2_cols - 1],
            'fill':       {'color': ref_biofuel_2['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True},
            'gap':        100
        })
    
    ref_worksheet15.insert_chart('R3', ref_tpes_biofuel_chart1)

    ##############################################################

    # TPES by fuel

    # access the sheet for production created above
    netz_worksheet16 = writer.sheets['Other fuels CN']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet16.set_column(2, netz_coal_1_cols + 1, None, space_format)
    netz_worksheet16.set_row(chart_height, None, header_format)
    netz_worksheet16.set_row(chart_height + netz_nuke_1_rows + 3, None, header_format)
    netz_worksheet16.set_row(chart_height + netz_nuke_1_rows + netz_biomass_1_rows + 6, None, header_format)
    netz_worksheet16.write(0, 0, economy + ' TPES nuclear, biomass and biofuels Carbon Neutrality', cell_format1)
    netz_worksheet16.write(1, 0, 'Units: Petajoules', cell_format2)
    
    # Create a TPES nuclear  chart
    if netz_nuke_1_rows > 0:
        netz_tpes_nuke_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_nuke_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_nuke_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_nuke_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_nuke_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Nuclear (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_nuke_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_nuke_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in netz_nuke_1['item_code_new'].unique():
            i = netz_nuke_1[netz_nuke_1['item_code_new'] == component].index[0]
            netz_tpes_nuke_chart1.add_series({
                'name':       ['Other fuels CN', chart_height + i + 1, 1],
                'categories': ['Other fuels CN', chart_height, 2, chart_height, netz_nuke_1_cols - 1],
                'values':     ['Other fuels CN', chart_height + i + 1, 2, chart_height + i + 1, netz_nuke_1_cols - 1],
                'fill':       {'color': netz_nuke_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        netz_worksheet16.insert_chart('B3', netz_tpes_nuke_chart1)

    else:
        pass

    # Create a TPES biomass chart
    netz_tpes_biomass_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_tpes_biomass_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_biomass_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_biomass_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_biomass_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Biomass (PJ)',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_biomass_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_tpes_biomass_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in fuel_final_nobunk[:-1]:
        i = netz_biomass_1[netz_biomass_1['item_code_new'] == component].index[0]
        netz_tpes_biomass_chart1.add_series({
            'name':       ['Other fuels CN', chart_height + netz_nuke_1_rows + i + 4, 1],
            'categories': ['Other fuels CN', chart_height + netz_nuke_1_rows + 3, 2,\
                chart_height + netz_nuke_1_rows + 3, netz_biomass_1_cols - 1],
            'values':     ['Other fuels CN', chart_height + netz_nuke_1_rows + i + 4, 2,\
                chart_height + netz_nuke_1_rows + i + 4, netz_biomass_1_cols - 1],
            'fill':       {'color': netz_biomass_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True},
            'gap':        100
        })
    
    netz_worksheet16.insert_chart('J3', netz_tpes_biomass_chart1)

    # Create a TPES biofuel chart
    netz_tpes_biofuel_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_tpes_biofuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_biofuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_biofuel_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_biofuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Biofuels (PJ)',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_biofuel_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    netz_tpes_biofuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in fuel_final_bunk[:-1]:
        i = netz_biofuel_2[netz_biofuel_2['item_code_new'] == component].index[0]
        netz_tpes_biofuel_chart1.add_series({
            'name':       ['Other fuels CN', chart_height + netz_nuke_1_rows + netz_biomass_1_rows + i + 7, 1],
            'categories': ['Other fuels CN', chart_height + netz_nuke_1_rows + netz_biomass_1_rows + 6, 2,\
                chart_height + netz_nuke_1_rows + netz_biomass_1_rows + 6, netz_biofuel_2_cols - 1],
            'values':     ['Other fuels CN', chart_height + netz_nuke_1_rows + netz_biomass_1_rows + i + 7, 2,\
                chart_height + netz_nuke_1_rows + netz_biomass_1_rows + i + 7, netz_biofuel_2_cols - 1],
            'fill':       {'color': netz_biofuel_2['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True},
            'gap':        100
        })
    
    netz_worksheet16.insert_chart('R3', netz_tpes_biofuel_chart1)

    #############################################################################################

    # FUEL consumptions and supply sheet

    # Access the workbook and second sheet with data from df2
    ref_worksheet41 = writer.sheets['Coal']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet41.set_column(1, ref_coalcons_1_cols + 1, None, space_format)
    ref_worksheet41.set_row(chart_height, None, header_format)
    ref_worksheet41.set_row(chart_height + ref_coalcons_1_rows + 3, None, header_format)
    ref_worksheet41.set_row((2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + 6, None, header_format)
    ref_worksheet41.set_row((2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + netz_coalcons_1_rows + 9, None, header_format)
    ref_worksheet41.write(0, 0, economy + ' coal Reference', cell_format1)
    ref_worksheet41.write(chart_height + ref_coalcons_1_rows + ref_coal_1_rows + 6, 0, economy + ' coal Carbon Neutrality', cell_format1)
    ref_worksheet41.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a FED sector area chart
    if ref_coalcons_1_rows > 0:
        ref_coalcons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_coalcons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_coalcons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_coalcons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_coalcons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_coalcons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_coalcons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_coalcons_1_rows):
            if not ref_coalcons_1['item_code_new'].iloc[i] in ['Total']:
                ref_coalcons_chart1.add_series({
                    'name':       ['Coal', chart_height + i + 1, 1],
                    'categories': ['Coal', chart_height, 2, chart_height, ref_coalcons_1_cols - 1],
                    'values':     ['Coal', chart_height + i + 1, 2, chart_height + i + 1, ref_coalcons_1_cols - 1],
                    'fill':       {'color': ref_coalcons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet41.insert_chart('B3', ref_coalcons_chart1)
    
    else:
        pass

    # Create a TPES coal chart
    if ref_coal_1_rows > 0:
        ref_tpes_coal_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_coal_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_coal_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_coal_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_coal_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Coal (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_coal_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_coal_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_nobunk[:-1]:
            i = ref_coal_1[ref_coal_1['item_code_new'] == component].index[0]
            ref_tpes_coal_chart1.add_series({
                'name':       ['Coal', chart_height + ref_coalcons_1_rows + i + 4, 1],
                'categories': ['Coal', chart_height + ref_coalcons_1_rows + 3, 2,\
                    chart_height + ref_coalcons_1_rows + 3, ref_coal_1_cols - 1],
                'values':     ['Coal', chart_height + ref_coalcons_1_rows + i + 4, 2,\
                    chart_height + ref_coalcons_1_rows + i + 4, ref_coal_1_cols - 1],
                'fill':       {'color': ref_coal_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet41.insert_chart('J3', ref_tpes_coal_chart1)

    else:
        pass

    # Carbon Neutrality coal charts

    # Create a FED sector area chart
    if netz_coalcons_1_rows > 0:
        netz_coalcons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_coalcons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_coalcons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_coalcons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_coalcons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_coalcons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_coalcons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_coalcons_1_rows):
            if not netz_coalcons_1['item_code_new'].iloc[i] in ['Total']:
                netz_coalcons_chart1.add_series({
                    'name':       ['Coal', (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + i + 7, 1],
                    'categories': ['Coal', (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + 6, 2,\
                        (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + 6, netz_coalcons_1_cols - 1],
                    'values':     ['Coal', (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + i + 7, 2,\
                        (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + i + 7, netz_coalcons_1_cols - 1],
                    'fill':       {'color': netz_coalcons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet41.insert_chart('B' + str(chart_height + ref_coalcons_1_rows + ref_coal_1_rows + 9), netz_coalcons_chart1)

    else:
        pass

    # Create a TPES coal chart
    if netz_coal_1_rows > 0:
        netz_tpes_coal_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_coal_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_coal_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_coal_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_coal_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Coal (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_coal_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_coal_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_nobunk[:-1]:
            i = netz_coal_1[netz_coal_1['item_code_new'] == component].index[0]
            netz_tpes_coal_chart1.add_series({
                'name':       ['Coal', (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + netz_coalcons_1_rows + i + 10, 1],
                'categories': ['Coal', (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + netz_coalcons_1_rows + 9, 2,\
                    (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + netz_coalcons_1_rows + 9, netz_coal_1_cols - 1],
                'values':     ['Coal', (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + netz_coalcons_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_coalcons_1_rows + ref_coal_1_rows + netz_coalcons_1_rows + i + 10, netz_coal_1_cols - 1],
                'fill':       {'color': netz_coal_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet41.insert_chart('J' + str(chart_height + ref_coalcons_1_rows + ref_coal_1_rows + 9), netz_tpes_coal_chart1)

    else:
        pass

    ##############
    # Natural gas
    
    # Access the workbook and second sheet with data from df2
    ref_worksheet42 = writer.sheets['Gas']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet42.set_column(1, ref_gascons_1_cols + 1, None, space_format)
    ref_worksheet42.set_row(chart_height, None, header_format)
    ref_worksheet42.set_row(chart_height + ref_gascons_1_rows + 3, None, header_format)
    ref_worksheet42.set_row((2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + 6, None, header_format)
    ref_worksheet42.set_row((2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + netz_gascons_1_rows + 9, None, header_format)
    ref_worksheet42.write(0, 0, economy + ' gas Reference', cell_format1)
    ref_worksheet42.write(chart_height + ref_gascons_1_rows + ref_gas_1_rows + 6, 0, economy + ' gas Carbon Neutrality', cell_format1)
    ref_worksheet42.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a FED sector area chart
    if ref_gascons_1_rows > 0:
        ref_gascons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_gascons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_gascons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_gascons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_gascons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_gascons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_gascons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_gascons_1_rows):
            if not ref_gascons_1['item_code_new'].iloc[i] in ['Total']:
                ref_gascons_chart1.add_series({
                    'name':       ['Gas', chart_height + i + 1, 1],
                    'categories': ['Gas', chart_height, 2, chart_height, ref_gascons_1_cols - 1],
                    'values':     ['Gas', chart_height + i + 1, 2, chart_height + i + 1, ref_gascons_1_cols - 1],
                    'fill':       {'color': ref_gascons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet42.insert_chart('B3', ref_gascons_chart1)

    else:
        pass

    # Create a TPES gas chart
    if ref_gas_1_rows > 0:
        ref_tpes_gas_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_gas_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_gas_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_gas_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_gas_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Gas (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_gas_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_gas_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_bunk[:-1]:
            i = ref_gas_1[ref_gas_1['item_code_new'] == component].index[0]
            ref_tpes_gas_chart1.add_series({
                'name':       ['Gas', chart_height + ref_gascons_1_rows + i + 4, 1],
                'categories': ['Gas', chart_height + ref_gascons_1_rows + 3, 2,\
                    chart_height + ref_gascons_1_rows + 3, ref_gas_1_cols - 1],
                'values':     ['Gas', chart_height + ref_gascons_1_rows + i + 4, 2,\
                    chart_height + ref_gascons_1_rows + i + 4, ref_gas_1_cols - 1],
                'fill':       {'color': ref_gas_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet42.insert_chart('J3', ref_tpes_gas_chart1)

    else:
        pass

    # Carbon Neutrality 

    # Create a FED sector area chart
    if netz_gascons_1_rows > 0:
        netz_gascons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_gascons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_gascons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_gascons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_gascons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_gascons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_gascons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_gascons_1_rows):
            if not netz_gascons_1['item_code_new'].iloc[i] in ['Total']:
                netz_gascons_chart1.add_series({
                    'name':       ['Gas', (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + i + 7, 1],
                    'categories': ['Gas', (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + 6, 2,\
                        (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + 6, netz_gascons_1_cols - 1],
                    'values':     ['Gas', (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + i + 7, 2,\
                        (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + i + 7, netz_gascons_1_cols - 1],
                    'fill':       {'color': netz_gascons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet42.insert_chart('B' + str(chart_height + ref_gascons_1_rows + ref_gas_1_rows + 9), netz_gascons_chart1)

    else:
        pass

    # Create a TPES gas chart
    if netz_gas_1_rows > 0:
        netz_tpes_gas_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_gas_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_gas_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_gas_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_gas_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Gas (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_gas_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_gas_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_bunk[:-1]:
            i = netz_gas_1[netz_gas_1['item_code_new'] == component].index[0]
            netz_tpes_gas_chart1.add_series({
                'name':       ['Gas', (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + netz_gascons_1_rows + i + 10, 1],
                'categories': ['Gas', (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + netz_gascons_1_rows + 9, 2,\
                    (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + netz_gascons_1_rows + 9, netz_gas_1_cols - 1],
                'values':     ['Gas', (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + netz_gascons_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_gascons_1_rows + ref_gas_1_rows + netz_gascons_1_rows + i + 10, netz_gas_1_cols - 1],
                'fill':       {'color': netz_gas_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet42.insert_chart('J' + str(chart_height + ref_gascons_1_rows + ref_gas_1_rows + 9), netz_tpes_gas_chart1)

    else:
        pass

    ##############
    # Crude
    
    # Access the workbook and second sheet with data from df2
    ref_worksheet43 = writer.sheets['Crude & NGL']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet43.set_column(1, ref_crudecons_1_cols + 1, None, space_format)
    ref_worksheet43.set_row(chart_height, None, header_format)
    ref_worksheet43.set_row(chart_height + ref_crudecons_1_rows + 3, None, header_format)
    ref_worksheet43.set_row((2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + 6, None, header_format)
    ref_worksheet43.set_row((2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + netz_crudecons_1_rows + 9, None, header_format)
    ref_worksheet43.write(0, 0, economy + ' crude & NGL Reference', cell_format1)
    ref_worksheet43.write(chart_height + ref_crudecons_1_rows + ref_crude_1_rows + 6, 0, economy + ' crude & NGL Carbon Neutrality', cell_format1)
    ref_worksheet43.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a FED sector area chart
    if ref_crudecons_1_rows > 0:
        ref_crudecons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_crudecons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_crudecons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_crudecons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_crudecons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_crudecons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_crudecons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_crudecons_1_rows):
            if not ref_crudecons_1['item_code_new'].iloc[i] in ['Total']:
                ref_crudecons_chart1.add_series({
                    'name':       ['Crude & NGL', chart_height + i + 1, 1],
                    'categories': ['Crude & NGL', chart_height, 2, chart_height, ref_crudecons_1_cols - 1],
                    'values':     ['Crude & NGL', chart_height + i + 1, 2, chart_height + i + 1, ref_crudecons_1_cols - 1],
                    'fill':       {'color': ref_crudecons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet43.insert_chart('B3', ref_crudecons_chart1)

    else:
        pass

    # Create a TPES crude oil and NGL chart
    if ref_crude_1_rows > 0:
        ref_tpes_crude_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_crude_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_crude_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_crude_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_crude_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Crude oil & NGL (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_crude_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_crude_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_nobunk[:-1]:
            i = ref_crude_1[ref_crude_1['item_code_new'] == component].index[0]
            ref_tpes_crude_chart1.add_series({
                'name':       ['Crude & NGL', chart_height + ref_crudecons_1_rows + i + 4, 1],
                'categories': ['Crude & NGL', chart_height + ref_crudecons_1_rows + 3, 2,\
                    chart_height + ref_crudecons_1_rows + 3, ref_crude_1_cols - 1],
                'values':     ['Crude & NGL', chart_height + ref_crudecons_1_rows + i + 4, 2,\
                    chart_height + ref_crudecons_1_rows + i + 4, ref_crude_1_cols - 1],
                'fill':       {'color': ref_crude_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet43.insert_chart('J3', ref_tpes_crude_chart1)

    else:
        pass

    # Carbon Neutrality

    # Create a FED sector area chart
    if netz_crudecons_1_rows > 0:
        netz_crudecons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_crudecons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_crudecons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_crudecons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_crudecons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_crudecons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_crudecons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_crudecons_1_rows):
            if not netz_crudecons_1['item_code_new'].iloc[i] in ['Total']:
                netz_crudecons_chart1.add_series({
                    'name':       ['Crude & NGL', (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + i + 7, 1],
                    'categories': ['Crude & NGL', (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + 6, 2,\
                        (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + 6, netz_crudecons_1_cols - 1],
                    'values':     ['Crude & NGL', (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + i + 7, 2,\
                        (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + i + 7, netz_crudecons_1_cols - 1],
                    'fill':       {'color': netz_crudecons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet43.insert_chart('B' + str(chart_height + ref_crudecons_1_rows + ref_crude_1_rows + 9), netz_crudecons_chart1)

    else:
        pass

    # Create a TPES gas chart
    if netz_crude_1_rows > 0:
        netz_tpes_crude_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_crude_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_crude_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_crude_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_crude_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            #'name': 'Crude oil & NGL (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_crude_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_crude_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_nobunk[:-1]:
            i = netz_crude_1[netz_crude_1['item_code_new'] == component].index[0]
            netz_tpes_crude_chart1.add_series({
                'name':       ['Crude & NGL', (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + netz_crudecons_1_rows + i + 10, 1],
                'categories': ['Crude & NGL', (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + netz_crudecons_1_rows + 9, 2,\
                    (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + netz_crudecons_1_rows + 9, netz_crude_1_cols - 1],
                'values':     ['Crude & NGL', (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + netz_crudecons_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_crudecons_1_rows + ref_crude_1_rows + netz_crudecons_1_rows + i + 10, netz_crude_1_cols - 1],
                'fill':       {'color': netz_crude_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet43.insert_chart('J' + str(chart_height + ref_crudecons_1_rows + ref_crude_1_rows + 9), netz_tpes_crude_chart1)

    else:
        pass

    ##############
    # Petroleum products
    
    # Access the workbook and second sheet with data from df2
    ref_worksheet44 = writer.sheets['Petroleum products']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet44.set_column(1, ref_petprodcons_1_cols + 1, None, space_format)
    ref_worksheet44.set_row(chart_height, None, header_format)
    ref_worksheet44.set_row(chart_height + ref_petprodcons_1_rows + 3, None, header_format)
    ref_worksheet44.set_row((2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + 6, None, header_format)
    ref_worksheet44.set_row((2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + netz_petprodcons_1_rows + 9, None, header_format)
    ref_worksheet44.write(0, 0, economy + ' petroleum products Reference', cell_format1)
    ref_worksheet44.write(chart_height + ref_petprodcons_1_rows + ref_petprod_2_rows + 6, 0, economy + ' petroleum products Carbon Neutrality', cell_format1)
    ref_worksheet44.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a FED sector area chart
    if ref_petprodcons_1_rows > 0:
        ref_petprodcons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_petprodcons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_petprodcons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_petprodcons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_petprodcons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_petprodcons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_petprodcons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_petprodcons_1_rows):
            if not ref_petprodcons_1['item_code_new'].iloc[i] in ['Total']:
                ref_petprodcons_chart1.add_series({
                    'name':       ['Petroleum products', chart_height + i + 1, 1],
                    'categories': ['Petroleum products', chart_height, 2, chart_height, ref_petprodcons_1_cols - 1],
                    'values':     ['Petroleum products', chart_height + i + 1, 2, chart_height + i + 1, ref_petprodcons_1_cols - 1],
                    'fill':       {'color': ref_petprodcons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet44.insert_chart('B3', ref_petprodcons_chart1)

    else:
        pass

    # Create a TPES petroleum products chart
    if ref_petprod_2_rows > 0:
        ref_tpes_petprod_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_petprod_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_petprod_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_petprod_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_petprod_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Petroleum products (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_petprod_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_petprod_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_ref[:-1]:
            i = ref_petprod_2[ref_petprod_2['item_code_new'] == component].index[0]
            ref_tpes_petprod_chart1.add_series({
                'name':       ['Petroleum products', chart_height + ref_petprodcons_1_rows + i + 4, 1],
                'categories': ['Petroleum products', chart_height + ref_petprodcons_1_rows + 3, 2,\
                    chart_height + ref_petprodcons_1_rows + 3, ref_petprod_2_cols - 1],
                'values':     ['Petroleum products', chart_height + ref_petprodcons_1_rows + i + 4, 2,\
                    chart_height + ref_petprodcons_1_rows + i + 4, ref_petprod_2_cols - 1],
                'fill':       {'color': ref_petprod_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet44.insert_chart('J3', ref_tpes_petprod_chart1)

    else:
        pass

    # Carbon Neutrality
    
    # Create a FED sector area chart
    if netz_petprodcons_1_rows > 0:
        netz_petprodcons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_petprodcons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_petprodcons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_petprodcons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_petprodcons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_petprodcons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_petprodcons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_petprodcons_1_rows):
            if not netz_petprodcons_1['item_code_new'].iloc[i] in ['Total']:
                netz_petprodcons_chart1.add_series({
                    'name':       ['Petroleum products', (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + i + 7, 1],
                    'categories': ['Petroleum products', (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + 6, 2,\
                        (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + 6, netz_petprodcons_1_cols - 1],
                    'values':     ['Petroleum products', (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + i + 7, netz_petprodcons_1_cols - 1],
                    'fill':       {'color': netz_petprodcons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet44.insert_chart('B' + str(chart_height + ref_petprodcons_1_rows + ref_petprod_2_rows + 9), netz_petprodcons_chart1)

    else:
        pass

    # Create a TPES petroleum products chart
    if netz_petprod_2_rows > 0:
        netz_tpes_petprod_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_petprod_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_petprod_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_petprod_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_petprod_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Petroleum products (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_petprod_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_petprod_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_ref[:-1]:
            i = netz_petprod_2[netz_petprod_2['item_code_new'] == component].index[0]
            netz_tpes_petprod_chart1.add_series({
                'name':       ['Petroleum products', (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + netz_petprodcons_1_rows + i + 10, 1],
                'categories': ['Petroleum products', (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + netz_petprodcons_1_rows + 9, 2,\
                    (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + netz_petprodcons_1_rows + 9, netz_petprod_2_cols - 1],
                'values':     ['Petroleum products', (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + netz_petprodcons_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_petprodcons_1_rows + ref_petprod_2_rows + netz_petprodcons_1_rows + i + 10, netz_petprod_2_cols - 1],
                'fill':       {'color': netz_petprod_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet44.insert_chart('J' + str(chart_height + ref_petprodcons_1_rows + ref_petprod_2_rows + 9), netz_tpes_petprod_chart1)

    else:
        pass

    ##############
    # Hydrogen
    
    # Access the workbook and second sheet with data from df2
    ref_worksheet45 = writer.sheets['Hydrogen']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet45.set_column(1, ref_hyd_1_cols + 1, None, space_format)
    ref_worksheet45.set_row(chart_height, None, header_format)
    ref_worksheet45.set_row(chart_height + ref_hyd_1_rows + 3, None, header_format)
    ref_worksheet45.set_row(chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + 6, None, header_format)
    ref_worksheet45.set_row((2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + 9, None, header_format)
    ref_worksheet45.set_row((2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + netz_hyd_1_rows + 12, None, header_format)
    ref_worksheet45.set_row((2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + netz_hyd_1_rows + netz_hydrogen_3_rows + 15, None, header_format)
    ref_worksheet45.write(0, 0, economy + ' hydrogen Reference', cell_format1)
    ref_worksheet45.write(chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + 9, 0, economy + ' hydrogen Carbon Neutrality', cell_format1)
    ref_worksheet45.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a FED sector area chart
    if ref_hyd_1_rows > 0:
        ref_hydrogen_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_hydrogen_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_hydrogen_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_hydrogen_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_hydrogen_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_hydrogen_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_hydrogen_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_hyd_1_rows):
            if not ref_hyd_1['item_code_new'].iloc[i] in ['Total']:
                ref_hydrogen_chart1.add_series({
                    'name':       ['Hydrogen', chart_height + i + 1, 1],
                    'categories': ['Hydrogen', chart_height, 2, chart_height, ref_hyd_1_cols - 1],
                    'values':     ['Hydrogen', chart_height + i + 1, 2, chart_height + i + 1, ref_hyd_1_cols - 1],
                    'fill':       {'color': ref_hyd_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet45.insert_chart('B3', ref_hydrogen_chart1)

    else:
        pass

    # Create a TPES hydrogen chart
    if ref_hydrogen_3_rows > 0:
        ref_tpes_hydrogen_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_hydrogen_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_hydrogen_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_hydrogen_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_hydrogen_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Hydrogen (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_hydrogen_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_hydrogen_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in ref_hydrogen_3['Technology'].unique():
            i = ref_hydrogen_3[ref_hydrogen_3['Technology'] == component].index[0]
            ref_tpes_hydrogen_chart1.add_series({
                'name':       ['Hydrogen', chart_height + ref_hyd_1_rows + i + 4, 1],
                'categories': ['Hydrogen', chart_height + ref_hyd_1_rows + 3, 2,\
                    chart_height + ref_hyd_1_rows + 3, ref_hydrogen_3_cols - 1],
                'values':     ['Hydrogen', chart_height + ref_hyd_1_rows + i + 4, 2,\
                    chart_height + ref_hyd_1_rows + i + 4, ref_hydrogen_3_cols - 1],
                'fill':       {'color': ref_hydrogen_3['Technology'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet45.insert_chart('J3', ref_tpes_hydrogen_chart1)

    else:
        pass

    # Create a TPES hydrogen chart
    if ref_hyd_use_1_rows > 0:
        ref_hyduse_chart1 = workbook.add_chart({'type': 'line'})
        ref_hyduse_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_hyduse_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_hyduse_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_hyduse_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Hydrogen (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_hyduse_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_hyduse_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in ref_hyd_use_1['FUEL'].unique():
            i = ref_hyd_use_1[ref_hyd_use_1['FUEL'] == component].index[0]
            if not ref_hyd_use_1['FUEL'].iloc[i] in ['Total']:
                ref_hyduse_chart1.add_series({
                    'name':       ['Hydrogen', chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + i + 7, 0],
                    'categories': ['Hydrogen', chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + 6, 2,\
                        chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + 6, ref_hyd_use_1_cols - 1],
                    'values':     ['Hydrogen', chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + i + 7, 2,\
                        chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + i + 7, ref_hyd_use_1_cols - 1],
                    'line':       {'color': ref_hyd_use_1['FUEL'].map(colours_dict).loc[i], 'width': 1.25}
                })

            else:
                pass
        
        ref_worksheet45.insert_chart('R3', ref_hyduse_chart1)

    else:
        pass

    # Carbon Neutrality
    
    # Create a FED sector area chart
    if netz_hyd_1_rows > 0:
        netz_hydrogen_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_hydrogen_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_hydrogen_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_hydrogen_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_hydrogen_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_hydrogen_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_hydrogen_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_hyd_1_rows):
            if not netz_hyd_1['item_code_new'].iloc[i] in ['Total']:
                netz_hydrogen_chart1.add_series({
                    'name':       ['Hydrogen', (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + i + 10, 1],
                    'categories': ['Hydrogen', (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + 9, 2,\
                        (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + 9, netz_hyd_1_cols - 1],
                    'values':     ['Hydrogen', (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + i + 10, 2,\
                        (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + i + 10, netz_hyd_1_cols - 1],
                    'fill':       {'color': netz_hyd_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass    
            
        ref_worksheet45.insert_chart('B' + str(chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + 12), netz_hydrogen_chart1)

    else:
        pass

    # Create a TPES hydrogen chart
    if  netz_hydrogen_3_rows > 0:
        netz_tpes_hydrogen_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_hydrogen_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_hydrogen_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_hydrogen_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_hydrogen_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Hydrogen (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_hydrogen_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_hydrogen_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in netz_hydrogen_3['Technology'].unique():
            i = netz_hydrogen_3[netz_hydrogen_3['Technology'] == component].index[0]
            netz_tpes_hydrogen_chart1.add_series({
                'name':       ['Hydrogen', (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + i + 13, 1],
                'categories': ['Hydrogen', (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + 12, 2,\
                    (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + 12, netz_hydrogen_3_cols - 1],
                'values':     ['Hydrogen', (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + i + 13, 2,\
                    (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + i + 13, netz_hydrogen_3_cols - 1],
                'fill':       {'color': netz_hydrogen_3['Technology'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet45.insert_chart('J' + str(chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + 12), netz_tpes_hydrogen_chart1)

    else:
        pass 

    # Create a TPES hydrogen chart
    if netz_hyd_use_1_rows > 0:
        netz_hyduse_chart1 = workbook.add_chart({'type': 'line'})
        netz_hyduse_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_hyduse_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_hyduse_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_hyduse_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Hydrogen (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_hyduse_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_hyduse_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in netz_hyd_use_1['FUEL'].unique():
            i = netz_hyd_use_1[netz_hyd_use_1['FUEL'] == component].index[0]
            if not netz_hyd_use_1['FUEL'].iloc[i] in ['Total']:
                netz_hyduse_chart1.add_series({
                    'name':       ['Hydrogen', (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + netz_hydrogen_3_rows + i + 16, 0],
                    'categories': ['Hydrogen', (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + netz_hydrogen_3_rows + 15, 2,\
                        (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + netz_hydrogen_3_rows + 15, netz_hyd_use_1_cols - 1],
                    'values':     ['Hydrogen', (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + netz_hydrogen_3_rows + i + 16, 2,\
                        (2 * chart_height) + ref_hyd_1_rows + ref_hydrogen_3_rows + netz_hyd_1_rows + ref_hyd_use_1_rows + netz_hydrogen_3_rows + i + 16, netz_hyd_use_1_cols - 1],
                    'line':       {'color': netz_hyd_use_1['FUEL'].map(colours_dict).loc[i], 'width': 1.25}
                })

            else:
                pass
        
        ref_worksheet45.insert_chart('R' + str(chart_height + ref_hyd_1_rows + ref_hydrogen_3_rows + ref_hyd_use_1_rows + 12), netz_hyduse_chart1)

    else:
        pass   

    ##############
    # Liquid and solid renewables
    
    # Access the workbook and second sheet with data from df2
    ref_worksheet46 = writer.sheets['Renewable fuels']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet46.set_column(1, ref_renewcons_1_cols + 1, None, space_format)
    ref_worksheet46.set_row(chart_height, None, header_format)
    ref_worksheet46.set_row(chart_height + ref_renewcons_1_rows + 3, None, header_format)
    ref_worksheet46.set_row((2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + 6, None, header_format)
    ref_worksheet46.set_row((2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + netz_renewcons_1_rows + 9, None, header_format)
    ref_worksheet46.write(0, 0, economy + ' liquid and solid renewables Reference', cell_format1)
    ref_worksheet46.write(chart_height + ref_renewcons_1_rows + ref_renew_2_rows + 6, 0, economy + ' liquid and solid renewables Carbon Neutrality', cell_format1)
    ref_worksheet46.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a FED sector area chart
    if ref_renewcons_1_rows > 0:
        ref_renewcons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_renewcons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_renewcons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_renewcons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_renewcons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_renewcons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_renewcons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_renewcons_1_rows):
            if not ref_renewcons_1['item_code_new'].iloc[i] in ['Buildings (biomass)', 'Total']:
                ref_renewcons_chart1.add_series({
                    'name':       ['Renewable fuels', chart_height + i + 1, 1],
                    'categories': ['Renewable fuels', chart_height, 2, chart_height, ref_renewcons_1_cols - 1],
                    'values':     ['Renewable fuels', chart_height + i + 1, 2, chart_height + i + 1, ref_renewcons_1_cols - 1],
                    'fill':       {'color': ref_renewcons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                if not ref_renewcons_1['item_code_new'].iloc[i] in ['Total']:
                    ref_renewcons_chart1.add_series({
                        'name':       ['Renewable fuels', chart_height + i + 1, 1],
                        'categories': ['Renewable fuels', chart_height, 2, chart_height, ref_renewcons_1_cols - 1],
                        'values':     ['Renewable fuels', chart_height + i + 1, 2, chart_height + i + 1, ref_renewcons_1_cols - 1],
                        'pattern':    {'bg_color': ref_renewcons_1['item_code_new'].map(colours_dict).loc[i],
                                    'fg_color': 'white',
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True}
                    })

                else:
                    pass

        ref_worksheet46.insert_chart('B3', ref_renewcons_chart1)

    else:
        pass

    # Create a TPES petroleum products chart
    if ref_renew_2_rows > 0:
        ref_tpes_renew_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_renew_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_renew_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_renew_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_renew_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Petroleum products (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_renew_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_renew_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_bunk[:-1]:
            i = ref_renew_2[ref_renew_2['item_code_new'] == component].index[0]
            ref_tpes_renew_chart1.add_series({
                'name':       ['Renewable fuels', chart_height + ref_renewcons_1_rows + i + 4, 1],
                'categories': ['Renewable fuels', chart_height + ref_renewcons_1_rows + 3, 2,\
                    chart_height + ref_renewcons_1_rows + 3, ref_renew_2_cols - 1],
                'values':     ['Renewable fuels', chart_height + ref_renewcons_1_rows + i + 4, 2,\
                    chart_height + ref_renewcons_1_rows + i + 4, ref_renew_2_cols - 1],
                'fill':       {'color': ref_renew_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet46.insert_chart('J3', ref_tpes_renew_chart1)

    else:
        pass

    # Carbon Neutrality
    
    # Create a FED sector area chart
    if netz_renewcons_1_rows > 0:
        netz_renewcons_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_renewcons_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_renewcons_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_renewcons_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_renewcons_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_renewcons_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_renewcons_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_renewcons_1_rows):
            if not netz_renewcons_1['item_code_new'].iloc[i] in ['Buildings (biomass)', 'Total']:
                netz_renewcons_chart1.add_series({
                    'name':       ['Renewable fuels', (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + i + 7, 1],
                    'categories': ['Renewable fuels', (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + 6, 2,\
                        (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + 6, netz_renewcons_1_cols - 1],
                    'values':     ['Renewable fuels', (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + i + 7, netz_renewcons_1_cols - 1],
                    'fill':       {'color': netz_renewcons_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                if not netz_renewcons_1['item_code_new'].iloc[i] in ['Total']:
                    netz_renewcons_chart1.add_series({
                        'name':       ['Renewable fuels', (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + i + 7, 1],
                        'categories': ['Renewable fuels', (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + 6, 2,\
                            (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + 6, netz_renewcons_1_cols - 1],
                        'values':     ['Renewable fuels', (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + i + 7, 2,\
                            (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + i + 7, netz_renewcons_1_cols - 1],
                        'pattern':    {'bg_color': netz_renewcons_1['item_code_new'].map(colours_dict).loc[i],
                                    'fg_color': 'white',
                                    'pattern': 'wide_downward_diagonal'},
                        'border':     {'none': True}
                    })

                else:
                    pass
            
        ref_worksheet46.insert_chart('B' + str(chart_height + ref_renewcons_1_rows + ref_renew_2_rows + 9), netz_renewcons_chart1)

    else:
        pass

    # Create a TPES petroleum products chart
    if netz_renew_2_rows > 0:
        netz_tpes_renew_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_renew_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_renew_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_renew_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_renew_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Petroleum products (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_renew_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_renew_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_bunk[:-1]:
            i = netz_renew_2[netz_renew_2['item_code_new'] == component].index[0]
            netz_tpes_renew_chart1.add_series({
                'name':       ['Renewable fuels', (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + netz_renewcons_1_rows + i + 10, 1],
                'categories': ['Renewable fuels', (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + netz_renewcons_1_rows + 9, 2,\
                    (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + netz_renewcons_1_rows + 9, netz_renew_2_cols - 1],
                'values':     ['Renewable fuels', (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + netz_renewcons_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_renewcons_1_rows + ref_renew_2_rows + netz_renewcons_1_rows + i + 10, netz_renew_2_cols - 1],
                'fill':       {'color': netz_renew_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        ref_worksheet46.insert_chart('J' + str(chart_height + ref_renewcons_1_rows + ref_renew_2_rows + 9), netz_tpes_renew_chart1)

    else:
        pass

    ############################################################################################################################################
    #VERSION 2: ONLY 1 BUILDINGS CATEGORY

    ##############
    # Liquid and solid renewables
    
    # Access the workbook and second sheet with data from df2
    both_worksheet61 = writer.sheets['Renewable fuels VER2']
        
    # Apply comma format and header format to relevant data rows
    both_worksheet61.set_column(1, ref_renewcons_2_cols + 1, None, space_format)
    both_worksheet61.set_row(chart_height, None, header_format)
    both_worksheet61.set_row(chart_height + ref_renewcons_2_rows + 3, None, header_format)
    both_worksheet61.set_row((2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + 6, None, header_format)
    both_worksheet61.set_row((2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + netz_renewcons_2_rows + 9, None, header_format)
    both_worksheet61.write(0, 0, economy + ' liquid and solid renewables Reference (NB: ONLY 1 BUILDINGS CATEGORY)', cell_format1)
    both_worksheet61.write(chart_height + ref_renewcons_2_rows + ref_renew_2_rows + 6, 0, economy + ' liquid and solid renewables Carbon Neutrality (NB: ONLY 1 BUILDINGS CATEGORY)', cell_format1)
    both_worksheet61.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a FED sector area chart
    if ref_renewcons_2_rows > 0:
        ref_renewcons_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_renewcons_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_renewcons_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        ref_renewcons_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_renewcons_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_renewcons_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_renewcons_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_renewcons_2_rows):
            if not ref_renewcons_2['item_code_new'].iloc[i] in ['Total']:
                ref_renewcons_chart2.add_series({
                    'name':       ['Renewable fuels VER2', chart_height + i + 1, 1],
                    'categories': ['Renewable fuels VER2', chart_height, 2, chart_height, ref_renewcons_2_cols - 1],
                    'values':     ['Renewable fuels VER2', chart_height + i + 1, 2, chart_height + i + 1, ref_renewcons_2_cols - 1],
                    'fill':       {'color': ref_renewcons_2['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass

        both_worksheet61.insert_chart('B3', ref_renewcons_chart2)

    else:
        pass

    # Create a TPES petroleum products chart
    if ref_renew_2_rows > 0:
        ref_tpes_renew_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_tpes_renew_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_tpes_renew_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        ref_tpes_renew_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_renew_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Petroleum products (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_tpes_renew_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_tpes_renew_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_bunk[:-1]:
            i = ref_renew_2[ref_renew_2['item_code_new'] == component].index[0]
            ref_tpes_renew_chart2.add_series({
                'name':       ['Renewable fuels VER2', chart_height + ref_renewcons_2_rows + i + 4, 1],
                'categories': ['Renewable fuels VER2', chart_height + ref_renewcons_2_rows + 3, 2,\
                    chart_height + ref_renewcons_2_rows + 3, ref_renew_2_cols - 1],
                'values':     ['Renewable fuels VER2', chart_height + ref_renewcons_2_rows + i + 4, 2,\
                    chart_height + ref_renewcons_2_rows + i + 4, ref_renew_2_cols - 1],
                'fill':       {'color': ref_renew_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        both_worksheet61.insert_chart('J3', ref_tpes_renew_chart2)

    else:
        pass

    # Carbon Neutrality
    
    # Create a FED sector area chart
    if netz_renewcons_2_rows > 0:
        netz_renewcons_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_renewcons_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_renewcons_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_renewcons_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_renewcons_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_renewcons_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_renewcons_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_renewcons_2_rows):
            if not netz_renewcons_2['item_code_new'].iloc[i] in ['Total']:
                netz_renewcons_chart2.add_series({
                    'name':       ['Renewable fuels VER2', (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + i + 7, 1],
                    'categories': ['Renewable fuels VER2', (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + 6, 2,\
                        (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + 6, netz_renewcons_2_cols - 1],
                    'values':     ['Renewable fuels VER2', (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + i + 7, 2,\
                        (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + i + 7, netz_renewcons_2_cols - 1],
                    'fill':       {'color': netz_renewcons_2['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass
            
        both_worksheet61.insert_chart('B' + str(chart_height + ref_renewcons_2_rows + ref_renew_2_rows + 9), netz_renewcons_chart2)

    else:
        pass

    # Create a TPES petroleum products chart
    if netz_renew_2_rows > 0:
        netz_tpes_renew_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_tpes_renew_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_tpes_renew_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_tpes_renew_chart2.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_renew_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Petroleum products (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_tpes_renew_chart2.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_tpes_renew_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for component in fuel_final_bunk[:-1]:
            i = netz_renew_2[netz_renew_2['item_code_new'] == component].index[0]
            netz_tpes_renew_chart2.add_series({
                'name':       ['Renewable fuels VER2', (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + netz_renewcons_2_rows + i + 10, 1],
                'categories': ['Renewable fuels VER2', (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + netz_renewcons_2_rows + 9, 2,\
                    (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + netz_renewcons_2_rows + 9, netz_renew_2_cols - 1],
                'values':     ['Renewable fuels VER2', (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + netz_renewcons_2_rows + i + 10, 2,\
                    (2 * chart_height) + ref_renewcons_2_rows + ref_renew_2_rows + netz_renewcons_2_rows + i + 10, netz_renew_2_cols - 1],
                'fill':       {'color': netz_renew_2['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True},
                'gap':        100
            })
        
        both_worksheet61.insert_chart('J' + str(chart_height + ref_renewcons_2_rows + ref_renew_2_rows + 9), netz_tpes_renew_chart2)

    else:
        pass

    ###############
    # Coal types

    # Access the workbook and second sheet with data from df2
    ref_worksheet47 = writer.sheets['Coal by type']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet47.set_column(1, ref_ct_prod1_cols + 1, None, space_format)
    ref_worksheet47.set_row(chart_height, None, header_format)
    ref_worksheet47.set_row(chart_height + ref_ct_prod1_rows + 3, None, header_format)
    ref_worksheet47.set_row(chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + 6, None, header_format)
    ref_worksheet47.set_row((2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + 9, None, header_format)
    ref_worksheet47.set_row((2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + 12, None, header_format)
    ref_worksheet47.set_row((2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + netz_ct_imports1_rows + 15, None, header_format)
    ref_worksheet47.write(0, 0, economy + ' coal types Reference', cell_format1)
    ref_worksheet47.write(chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + 9, 0, economy + ' coal types Carbon Neutrality', cell_format1)
    ref_worksheet47.write(1, 0, 'Units: Petajoules', cell_format2)

    # Coal production line chart 
    if ref_ct_prod1_rows > 0:
        ref_coalprod_chart1 = workbook.add_chart({'type': 'line'})
        ref_coalprod_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_coalprod_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_coalprod_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_coalprod_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_coalprod_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_coalprod_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_ct_prod1_rows):
            ref_coalprod_chart1.add_series({
                'name':       ['Coal by type', chart_height + i + 1, 0],
                'categories': ['Coal by type', chart_height, 2, chart_height, ref_ct_prod1_cols - 1],
                'values':     ['Coal by type', chart_height + i + 1, 2, chart_height + i + 1, ref_ct_prod1_cols - 1],
                'line':       {'color': ref_ct_prod1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet47.insert_chart('B3', ref_coalprod_chart1)

    else:
        pass

    # Coal imports line chart 
    if ref_ct_imports1_rows > 0:
        ref_coalim_chart1 = workbook.add_chart({'type': 'line'})
        ref_coalim_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_coalim_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_coalim_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_coalim_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_coalim_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_coalim_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_ct_imports1_rows):
            ref_coalim_chart1.add_series({
                'name':       ['Coal by type', chart_height + ref_ct_prod1_rows + i + 4, 0],
                'categories': ['Coal by type', chart_height + ref_ct_prod1_rows + 3, 2, chart_height + ref_ct_prod1_rows + 3, ref_ct_imports1_cols - 1],
                'values':     ['Coal by type', chart_height + ref_ct_prod1_rows + i + 4, 2, chart_height + ref_ct_prod1_rows + i + 4, ref_ct_imports1_cols - 1],
                'line':       {'color': ref_ct_imports1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet47.insert_chart('J3', ref_coalim_chart1)

    else:
        pass

    # Coal exports line chart 
    if ref_ct_exports1_rows > 0:
        ref_coalex_chart1 = workbook.add_chart({'type': 'line'})
        ref_coalex_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_coalex_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_coalex_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_coalex_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_coalex_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_coalex_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_ct_exports1_rows):
            ref_coalex_chart1.add_series({
                'name':       ['Coal by type', chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + i + 7, 0],
                'categories': ['Coal by type', chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + 6, 2,\
                    chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + 6, ref_ct_exports1_cols - 1],
                'values':     ['Coal by type', chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + i + 7, 2,\
                    chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + i + 7, ref_ct_exports1_cols - 1],
                'line':       {'color': ref_ct_exports1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet47.insert_chart('R3', ref_coalex_chart1)

    else:
        pass

    # Coal production line chart 
    if netz_ct_prod1_rows > 0:
        netz_coalprod_chart1 = workbook.add_chart({'type': 'line'})
        netz_coalprod_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_coalprod_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_coalprod_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_coalprod_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_coalprod_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_coalprod_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_ct_prod1_rows):
            netz_coalprod_chart1.add_series({
                'name':       ['Coal by type', (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + i + 10, 0],
                'categories': ['Coal by type', (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + 9, 2,\
                    (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + 9, netz_ct_prod1_cols - 1],
                'values':     ['Coal by type', (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + i + 10, netz_ct_prod1_cols - 1],
                'line':       {'color': netz_ct_prod1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet47.insert_chart('B' + str(chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + 12), netz_coalprod_chart1)

    else:
        pass

    # Coal imports line chart 
    if netz_ct_imports1_rows > 0:
        netz_coalim_chart1 = workbook.add_chart({'type': 'line'})
        netz_coalim_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_coalim_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_coalim_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_coalim_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_coalim_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_coalim_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_ct_imports1_rows):
            netz_coalim_chart1.add_series({
                'name':       ['Coal by type', (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + i + 13, 0],
                'categories': ['Coal by type', (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + 12, 2,\
                    (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + 12, netz_ct_imports1_cols - 1],
                'values':     ['Coal by type', (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + i + 13, 2,\
                    (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + i + 13, netz_ct_imports1_cols - 1],
                'line':       {'color': netz_ct_imports1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet47.insert_chart('J' + str(chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + 12), netz_coalim_chart1)

    else:
        pass

    # Coal exports line chart 
    if netz_ct_exports1_rows > 0:
        netz_coalex_chart1 = workbook.add_chart({'type': 'line'})
        netz_coalex_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_coalex_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_coalex_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_coalex_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_coalex_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_coalex_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_ct_exports1_rows):
            netz_coalex_chart1.add_series({
                'name':       ['Coal by type', (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + netz_ct_imports1_rows + i + 16, 0],
                'categories': ['Coal by type', (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + netz_ct_imports1_rows + 15, 2,\
                    (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + netz_ct_imports1_rows + 15, netz_ct_exports1_cols - 1],
                'values':     ['Coal by type', (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + netz_ct_imports1_rows + i + 16, 2,\
                    (2 * chart_height) + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + netz_ct_prod1_rows + netz_ct_imports1_rows + i + 16, netz_ct_exports1_cols - 1],
                'line':       {'color': netz_ct_exports1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet47.insert_chart('R' + str(chart_height + ref_ct_prod1_rows + ref_ct_imports1_rows + ref_ct_exports1_rows + 12), netz_coalex_chart1)

    else:
        pass

    ############################
    # Gas trade

    # Access the workbook and second sheet with data from df2
    ref_worksheet48 = writer.sheets['Gas trade']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet48.set_column(1, ref_gasim_1_cols + 1, None, space_format)
    ref_worksheet48.set_row(chart_height, None, header_format)
    ref_worksheet48.set_row(chart_height + ref_gasim_1_rows + 3, None, header_format)
    ref_worksheet48.set_row((2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + 6, None, header_format)
    ref_worksheet48.set_row((2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + netz_gasim_1_rows + 9, None, header_format)
    ref_worksheet48.write(0, 0, economy + ' gas trade Reference', cell_format1)
    ref_worksheet48.write(chart_height + ref_gasim_1_rows + ref_gasex_1_rows + 6, 0, economy + ' gas trade Carbon Neutrality', cell_format1)
    ref_worksheet48.write(1, 0, 'Units: Petajoules', cell_format2)

    # Gas imports line chart 
    if ref_gasim_1_rows > 0:
        ref_gasim_chart1 = workbook.add_chart({'type': 'line'})
        ref_gasim_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_gasim_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_gasim_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_gasim_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_gasim_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_gasim_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_gasim_1_rows):
            ref_gasim_chart1.add_series({
                'name':       ['Gas trade', chart_height + i + 1, 0],
                'categories': ['Gas trade', chart_height, 1, chart_height, ref_gasim_1_cols - 1],
                'values':     ['Gas trade', chart_height + i + 1, 1, chart_height + i + 1, ref_gasim_1_cols - 1],
                'line':       {'color': ref_gasim_1['Imports'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet48.insert_chart('B3', ref_gasim_chart1)

    else:
        pass

    # Gas exports line chart 
    if ref_gasex_1_rows > 0:
        ref_gasex_chart1 = workbook.add_chart({'type': 'line'})
        ref_gasex_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_gasex_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_gasex_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        ref_gasex_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_gasex_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_gasex_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_gasex_1_rows):
            ref_gasex_chart1.add_series({
                'name':       ['Gas trade', chart_height + ref_gasim_1_rows + i + 4, 0],
                'categories': ['Gas trade', chart_height + ref_gasim_1_rows + 3, 1,\
                    chart_height + ref_gasim_1_rows + 3, ref_gasex_1_cols - 1],
                'values':     ['Gas trade', chart_height + ref_gasim_1_rows + i + 4, 1,\
                    chart_height + ref_gasim_1_rows + i + 4, ref_gasex_1_cols - 1],
                'line':       {'color': ref_gasex_1['Exports'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet48.insert_chart('J3', ref_gasex_chart1)

    else:
        pass

    # Gas imports line chart 
    if netz_gasim_1_rows > 0:
        netz_gasim_chart1 = workbook.add_chart({'type': 'line'})
        netz_gasim_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_gasim_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_gasim_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_gasim_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_gasim_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_gasim_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_gasim_1_rows):
            netz_gasim_chart1.add_series({
                'name':       ['Gas trade', (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + i + 7, 0],
                'categories': ['Gas trade', (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + 6, 1,\
                    (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + 6, netz_gasim_1_cols - 1],
                'values':     ['Gas trade', (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + i + 7, 1,\
                    (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + i + 7, netz_gasim_1_cols - 1],
                'line':       {'color': netz_gasim_1['Imports'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet48.insert_chart('B' + str(chart_height + ref_gasim_1_rows + ref_gasex_1_rows + 9), netz_gasim_chart1)

    else:
        pass

    # Gas exports line chart 
    if netz_gasex_1_rows > 0:
        netz_gasex_chart1 = workbook.add_chart({'type': 'line'})
        netz_gasex_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_gasex_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_gasex_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 8,
            'line': {'color': '#bebebe'}
        })
            
        netz_gasex_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_gasex_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_gasex_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_gasex_1_rows):
            netz_gasex_chart1.add_series({
                'name':       ['Gas trade', (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + netz_gasim_1_rows + i + 10, 0],
                'categories': ['Gas trade', (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + netz_gasim_1_rows + 9, 1,\
                    (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + netz_gasim_1_rows + 9, netz_gasex_1_cols - 1],
                'values':     ['Gas trade', (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + netz_gasim_1_rows + i + 10, 1,\
                    (2 * chart_height) + ref_gasim_1_rows + ref_gasex_1_rows + netz_gasim_1_rows + i + 10, netz_gasex_1_cols - 1],
                'line':       {'color': netz_gasex_1['Exports'].map(colours_dict).loc[i], 'width': 1.25}
            })    
            
        ref_worksheet48.insert_chart('J' + str(chart_height + ref_gasim_1_rows + ref_gasex_1_rows + 9), netz_gasex_chart1)

    else:
        pass

    # Electricity charts

    # Access the workbook and second sheet with data from df2
    ref_worksheet49 = writer.sheets['Electricity']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet49.set_column(1, ref_elec_1_cols + 1, None, space_format)
    ref_worksheet49.set_row(chart_height, None, header_format)
    ref_worksheet49.set_row((2 * chart_height) + ref_elec_1_rows + 3, None, header_format)
    ref_worksheet49.write(0, 0, economy + ' electricity consumption Reference', cell_format1)
    ref_worksheet49.write(chart_height + ref_elec_1_rows + 3, 0, economy + ' electricity consumption Carbon Neutrality', cell_format1)
    ref_worksheet49.write(1, 0, 'Units: Petajoules', cell_format2)

    # Create a Electricity sector area chart
    # REFERENCE
    if ref_elec_1_rows > 0:
        ref_elecsec_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_elecsec_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_elecsec_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_elecsec_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        ref_elecsec_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        ref_elecsec_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        ref_elecsec_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_elec_1_rows):
            if not ref_elec_1['item_code_new'].iloc[i] in ['Total']:            
                ref_elecsec_chart1.add_series({
                    'name':       ['Electricity', chart_height + i + 1, 1],
                    'categories': ['Electricity', chart_height, 2, chart_height, ref_elec_1_cols - 1],
                    'values':     ['Electricity', chart_height + i + 1, 2, chart_height + i + 1, ref_elec_1_cols - 1],
                    'fill':       {'color': ref_elec_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass
            
        ref_worksheet49.insert_chart('B3', ref_elecsec_chart1)

    else:
        pass

    # Create a Electricity sector area chart
    # CARBON NEUTRALITY
    if netz_elec_1_rows > 0:
        netz_elecsec_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_elecsec_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_elecsec_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_elecsec_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'crossing': 19,
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'position_axis': 'on_tick',
            'interval_unit': 10,
            'line': {'color': '#bebebe'}
        })
            
        netz_elecsec_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'label_position': 'low',
            # 'name': 'PJ',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#323232',
                     'width': 1,
                     'dash_type': 'square_dot'}
        })
            
        netz_elecsec_chart1.set_legend({
            'font': {'name': 'Segoe UI', 'size': 9}
            #'none': True
        })
            
        netz_elecsec_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_elec_1_rows):
            if not netz_elec_1['item_code_new'].iloc[i] in ['Total']:            
                netz_elecsec_chart1.add_series({
                    'name':       ['Electricity', (2 * chart_height) + ref_elec_1_rows + i + 4, 1],
                    'categories': ['Electricity', (2 * chart_height) + ref_elec_1_rows + 3, 2,\
                        (2 * chart_height) + ref_elec_1_rows + 3, netz_elec_1_cols - 1],
                    'values':     ['Electricity', (2 * chart_height) + ref_elec_1_rows + i + 4, 2,\
                        (2 * chart_height) + ref_elec_1_rows + i + 4, netz_elec_1_cols - 1],
                    'fill':       {'color': netz_elec_1['item_code_new'].map(colours_dict).loc[i]},
                    'border':     {'none': True}
                })

            else:
                pass
            
        ref_worksheet49.insert_chart('B' + str(chart_height + ref_elec_1_rows + 6), netz_elecsec_chart1)

    else:
        pass

    ##############################################################
    # Carbon intensity chart

    # Access the workbook and second sheet
    both_worksheet38 = writer.sheets['CO2 intensity']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet38.set_column(2, ref_co2int_2_cols + 1, None, space_format)
    both_worksheet38.set_row(chart_height, None, header_format)
    both_worksheet38.set_row(chart_height + ref_co2int_2_rows + 3, None, header_format)
    both_worksheet38.set_row(chart_height + ref_co2int_2_rows + netz_co2int_2_rows + 6, None, header_format)
    both_worksheet38.write(0, 0, economy + ' carbon intensity and emissions', cell_format1)

    # line chart
    co2int_chart1 = workbook.add_chart({'type': 'line'})
    co2int_chart1.set_size({
        'width': 500,            
        'height': 300
    })
        
    co2int_chart1.set_chartarea({
        'border': {'none': True}
    })
        
    co2int_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
            
    co2int_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        'name': 'CO2 intensity',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0.00',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
            
    co2int_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
            
    co2int_chart1.set_title({
        'none': True
    })
        
    # Configure the series of the chart from the dataframe data.
    i = ref_co2int_2[ref_co2int_2['fuel_code'] == 'Reference'].index[0]
    co2int_chart1.add_series({
        'name':       ['CO2 intensity', chart_height + i + 1, 0],
        'categories': ['CO2 intensity', chart_height, 2,\
            chart_height, ref_co2int_2_cols - 1],
        'values':     ['CO2 intensity', chart_height + i + 1, 2,\
            chart_height + i + 1, ref_co2int_2_cols - 1],
        'line':       {'color': ref_co2int_2['fuel_code'].map(colours_dict).loc[i],
                       'width': 1.5}
        })

    j = netz_co2int_2[netz_co2int_2['fuel_code'] == 'Carbon neutrality'].index[0]
    co2int_chart1.add_series({
        'name':       ['CO2 intensity', chart_height + ref_co2int_2_rows + j + 4, 0],
        'categories': ['CO2 intensity', chart_height + ref_co2int_2_rows + 3, 2,\
            chart_height + ref_co2int_2_rows + 3, netz_co2int_2_cols - 1],
        'values':     ['CO2 intensity', chart_height + ref_co2int_2_rows + j + 4, 2,\
            chart_height + ref_co2int_2_rows + j + 4, netz_co2int_2_cols - 1],
        'line':       {'color': netz_co2int_2['fuel_code'].map(colours_dict).loc[j],
                       'width': 1.5}
        })
                    
    both_worksheet38.insert_chart('B3', co2int_chart1)

    # Create a Emissions line chart with higher level aggregation
    emiss_chart1 = workbook.add_chart({'type': 'line'})
    emiss_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    emiss_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    emiss_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    emiss_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        'name': 'Million tonnes CO2',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                 'width': 1,
                 'dash_type': 'square_dot'}
    })
        
    emiss_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    emiss_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(emiss_total_1_rows):
        emiss_chart1.add_series({
            'name':       ['CO2 intensity', chart_height + ref_co2int_2_rows + netz_co2int_2_rows + i + 7, 0],
            'categories': ['CO2 intensity', chart_height + ref_co2int_2_rows + netz_co2int_2_rows + 6, 2,\
                chart_height + ref_co2int_2_rows + netz_co2int_2_rows + 6, netz_co2int_2_cols - 1],
            'values':     ['CO2 intensity', chart_height + ref_co2int_2_rows + netz_co2int_2_rows + i + 7, 2,\
                chart_height + ref_co2int_2_rows + netz_co2int_2_rows + i + 7, netz_co2int_2_cols - 1],
            'line':       {'color': emiss_total_1['fuel_code'].map(colours_dict).loc[i], 
                            'width': 1.5}
        })    
        
    both_worksheet38.insert_chart('J3', emiss_chart1)

    # Emissions wedge charts

    # Access the workbook and second sheet
    both_worksheet39 = writer.sheets['CO2 wedge']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet39.set_column(2, emissions_wedge_1_cols + 1, None, space_format)
    both_worksheet39.set_row(chart_height, None, header_format)
    both_worksheet39.set_row(chart_height + emissions_wedge_1_rows + 3, None, header_format)
    both_worksheet39.write(0, 0, economy + ' emissions wedge charts', cell_format1)

    # Wedge chart: Sector

    emiss_wedge_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    emiss_wedge_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    emiss_wedge_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    emiss_wedge_chart1.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    emiss_wedge_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#828282',
                'width': 1,
                'dash_type': 'square_dot'}
    })

    emiss_wedge_chart1.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    emiss_wedge_chart1.set_title({
        'none': True
    })

    # line series' for adding in a line chart
    emiss_line_1 = workbook.add_chart({'type': 'line'})
    
    # Configure the series of the chart from the dataframe data.
    for i in list(emissions_wedge_1.loc[emissions_wedge_1['item_code_new'].isna()].index) +\
        list(emissions_wedge_1.loc[emissions_wedge_1['item_code_new'].isna() == False].index):
        if emissions_wedge_1['item_code_new'].iloc[i] in ['Power', 'Own use', 'Industry', 'Transport', 'Buildings', 'Agriculture', 'Non-specified']:
            emiss_wedge_chart1.add_series({
                'name':       ['CO2 wedge', chart_height + i + 1, 1],
                'categories': ['CO2 wedge', chart_height, 2, chart_height, emissions_wedge_1_cols - 1],
                'values':     ['CO2 wedge', chart_height + i + 1, 2, chart_height + i + 1, emissions_wedge_1_cols - 1],
                'fill':       {'color': emissions_wedge_1['item_code_new'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        elif emissions_wedge_1['item_code_new'].iloc[i] in [np.nan]:
            emiss_wedge_chart1.add_series({
                'name':       ['CO2 wedge', chart_height + i + 1, 1],
                'categories': ['CO2 wedge', chart_height, 2, chart_height, emissions_wedge_1_cols - 1],
                'values':     ['CO2 wedge', chart_height + i + 1, 2, chart_height + i + 1, emissions_wedge_1_cols - 1],
                'fill':       {'none': True},
                'border':     {'none': True}
            })

        elif emissions_wedge_1['item_code_new'].iloc[i] in ['Reference', 'Carbon neutrality']:
            emiss_line_1.add_series({
                'name':       ['CO2 wedge', chart_height + i + 1, 1],
                'categories': ['CO2 wedge', chart_height, 2, chart_height, emissions_wedge_1_cols - 1],
                'values':     ['CO2 wedge', chart_height + i + 1, 2, chart_height + i + 1, emissions_wedge_1_cols - 1],
                'line':       {'color': emissions_wedge_1['item_code_new'].map(colours_dict).loc[i],
                               'width': 1.25}
            })

        else:
            pass

    emiss_wedge_chart1.combine(emiss_line_1)
        
    both_worksheet39.insert_chart('B3', emiss_wedge_chart1)

    # Wedge chart: Fuel

    emiss_wedge_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    emiss_wedge_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    emiss_wedge_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    emiss_wedge_chart2.set_x_axis({
        # 'name': 'Year',
        'label_position': 'low',
        'crossing': 19,
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'position_axis': 'on_tick',
        'interval_unit': 10,
        'line': {'color': '#bebebe'}
    })
        
    emiss_wedge_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'label_position': 'low',
        # 'name': 'PJ',
        'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#323232',
                'width': 1,
                'dash_type': 'square_dot'}
    })
        
    emiss_wedge_chart2.set_legend({
        'font': {'name': 'Segoe UI', 'size': 9}
        #'none': True
    })
        
    emiss_wedge_chart2.set_title({
        'none': True
    })

    # line series' for adding in a line chart
    emiss_line_2 = workbook.add_chart({'type': 'line'})
    
    # Configure the series of the chart from the dataframe data.
    for i in list(emissions_wedge_2.loc[emissions_wedge_2['fuel_code'].isna()].index) +\
        list(emissions_wedge_2.loc[emissions_wedge_2['fuel_code'].isna() == False].index):
        if emissions_wedge_2['fuel_code'].iloc[i] in ['Coal', 'Oil', 'Gas', 'Heat & others']:
            emiss_wedge_chart2.add_series({
                'name':       ['CO2 wedge', chart_height + emissions_wedge_1_rows + i + 4, 0],
                'categories': ['CO2 wedge', chart_height + emissions_wedge_1_rows + 3, 2,\
                    chart_height + emissions_wedge_1_rows + 3, emissions_wedge_2_cols - 1],
                'values':     ['CO2 wedge', chart_height + emissions_wedge_1_rows + i + 4, 2,\
                    chart_height + emissions_wedge_1_rows + i + 4, emissions_wedge_2_cols - 1],
                'fill':       {'color': emissions_wedge_2['fuel_code'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        elif emissions_wedge_2['fuel_code'].iloc[i] in [np.nan]:
            emiss_wedge_chart2.add_series({
                'name':       ['CO2 wedge', chart_height + emissions_wedge_1_rows + i + 4, 0],
                'categories': ['CO2 wedge', chart_height + emissions_wedge_1_rows + 3, 2,\
                    chart_height + emissions_wedge_1_rows + 3, emissions_wedge_2_cols - 1],
                'values':     ['CO2 wedge', chart_height + emissions_wedge_1_rows + i + 4, 2,\
                    chart_height + emissions_wedge_1_rows + i + 4, emissions_wedge_2_cols - 1],
                'fill':       {'none': True},
                'border':     {'none': True}
            })

        elif emissions_wedge_2['fuel_code'].iloc[i] in ['Reference', 'Carbon neutrality']:
            emiss_line_2.add_series({
                'name':       ['CO2 wedge', chart_height + emissions_wedge_1_rows + i + 4, 0],
                'categories': ['CO2 wedge', chart_height + emissions_wedge_1_rows + 3, 2,\
                    chart_height + emissions_wedge_1_rows + 3, emissions_wedge_2_cols - 1],
                'values':     ['CO2 wedge', chart_height + emissions_wedge_1_rows + i + 4, 2,\
                    chart_height + emissions_wedge_1_rows + i + 4, emissions_wedge_2_cols - 1],
                'line':       {'color': emissions_wedge_2['fuel_code'].map(colours_dict).loc[i],
                               'width': 1.25}
            })

        else:
            pass

    emiss_wedge_chart2.combine(emiss_line_2)
        
    both_worksheet39.insert_chart('J3', emiss_wedge_chart2)

    # Kaya identity waterfall charts

    # Access the workbook and second sheet
    both_worksheet51 = writer.sheets['CO2 breakdown']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet51.set_column(1, ref_kaya_1_cols + 1, None, space_format)
    both_worksheet51.set_row(chart_height, None, header_format)
    both_worksheet51.set_row(chart_height + ref_kaya_1_rows + 3, None, header_format)
    both_worksheet51.write(0, 0, economy + ' Kaya waterfall charts (emissions deconstruction)', cell_format1)

    # First kaya waterfall
    if ref_kaya_1_rows > 0:
        ref_kaya_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_kaya_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_kaya_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_kaya_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            #'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_kaya_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Petroleum products (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        ref_kaya_chart1.set_legend({
            #'font': {'name': 'Segoe UI', 'size': 9}
            'none': True
        })
            
        ref_kaya_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_kaya_1_rows):
            if ref_kaya_1['Reference'].iloc[i] in ['initial']:
                ref_kaya_chart1.add_series({
                    'name':       ['CO2 breakdown', chart_height + i + 1, 0],
                    'categories': ['CO2 breakdown', chart_height, 1, chart_height, ref_kaya_1_cols - 1],
                    'values':     ['CO2 breakdown', chart_height + i + 1, 1, chart_height + i + 1, ref_kaya_1_cols - 1],
                    'fill':       {'color': ref_kaya_1['Reference'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        50
                })

            elif ref_kaya_1['Reference'].iloc[i] in ['improve', 'no improve']:
                ref_kaya_chart1.add_series({
                    'name':       ['CO2 breakdown', chart_height + i + 1, 0],
                    'categories': ['CO2 breakdown', chart_height, 1, chart_height, ref_kaya_1_cols - 1],
                    'values':     ['CO2 breakdown', chart_height + i + 1, 1, chart_height + i + 1, ref_kaya_1_cols - 1],
                    'fill':       {'color': ref_kaya_1['Reference'].map(colours_dict).loc[i],
                                'transparency': 50},
                    'border':     {'none': True},
                    'gap':        50
                })

            else:
                ref_kaya_chart1.add_series({
                    'name':       ['CO2 breakdown', chart_height + i + 1, 0],
                    'categories': ['CO2 breakdown', chart_height, 1, chart_height, ref_kaya_1_cols - 1],
                    'values':     ['CO2 breakdown', chart_height + i + 1, 1, chart_height + i + 1, ref_kaya_1_cols - 1],
                    'fill':       {'none': True},
                    'border':     {'none': True},
                    'gap':        50
                })
        
        both_worksheet51.insert_chart('B3', ref_kaya_chart1)

    else:
        pass

    # Second kaya waterfall
    if netz_kaya_1_rows > 0:
        netz_kaya_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_kaya_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_kaya_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_kaya_chart1.set_x_axis({
            # 'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            #'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_kaya_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            # 'name': 'Petroleum products (PJ)',
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_kaya_chart1.set_legend({
            #'font': {'name': 'Segoe UI', 'size': 9}
            'none': True
        })
            
        netz_kaya_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(netz_kaya_1_rows):
            if netz_kaya_1['Carbon neutrality'].iloc[i] in ['initial']:
                netz_kaya_chart1.add_series({
                    'name':       ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 0],
                    'categories': ['CO2 breakdown', chart_height + ref_kaya_1_rows + 3, 1, chart_height + ref_kaya_1_rows + 3, netz_kaya_1_cols - 1],
                    'values':     ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 1, chart_height + ref_kaya_1_rows + i + 4, netz_kaya_1_cols - 1],
                    'fill':       {'color': netz_kaya_1['Carbon neutrality'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        50
                })

            elif netz_kaya_1['Carbon neutrality'].iloc[i] in ['improve', 'no improve']:
                netz_kaya_chart1.add_series({
                    'name':       ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 0],
                    'categories': ['CO2 breakdown', chart_height + ref_kaya_1_rows + 3, 1, chart_height + ref_kaya_1_rows + 3, netz_kaya_1_cols - 1],
                    'values':     ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 1, chart_height + ref_kaya_1_rows + i + 4, netz_kaya_1_cols - 1],
                    'fill':       {'color': netz_kaya_1['Carbon neutrality'].map(colours_dict).loc[i],
                                   'transparency': 50},
                    'border':     {'none': True},
                    'gap':        50
                })

            else:
                netz_kaya_chart1.add_series({
                    'name':       ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 0],
                    'categories': ['CO2 breakdown', chart_height + ref_kaya_1_rows + 3, 1, chart_height + ref_kaya_1_rows + 3, netz_kaya_1_cols - 1],
                    'values':     ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 1, chart_height + ref_kaya_1_rows + i + 4, netz_kaya_1_cols - 1],
                    'fill':       {'none': True},
                    'border':     {'none': True},
                    'gap':        50
                })
        
        both_worksheet51.insert_chart('J3', netz_kaya_chart1)

    else:
        pass   

    writer.save()

print('Bling blang blaow, you have some charts now')

