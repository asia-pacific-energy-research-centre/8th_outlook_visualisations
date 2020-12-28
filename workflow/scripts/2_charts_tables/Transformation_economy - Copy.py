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

# Read in mapping file

Mapping_file = pd.read_excel(path_mapping + '/OSeMOSYS mapping.xlsx', sheet_name = 'Mapping',  skiprows = 1)

# Subset the mapping file so that it's just transformation

Map_trans = Mapping_file[Mapping_file['Balance'] == 'TRANS'].reset_index(drop = True)

# Define unique workbook and sheet combinations

Unique_trans = Map_trans.groupby(['Workbook', 'Sheet']).size().reset_index().loc[:, ['Workbook', 'Sheet']]

# Determine list of files to read based on the workbooks identified in the mapping file

file_trans = pd.DataFrame()

for i in range(len(Unique_trans['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in OSeMOSYS_filenames if Unique_trans['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_trans['Workbook'].unique()[i]})
    file_trans = file_trans.append(_file)

file_trans = file_trans.merge(Unique_trans, how = 'outer', on = 'Workbook')

# Create empty dataframe to store aggregated results 

aggregate_df1 = pd.DataFrame()

# Now read in the OSeMOSYS output files so that that they're all in one data frame (aggregate_df1)

for i in range(file_trans.shape[0]):
    _df = pd.read_excel(file_trans.iloc[i, 0], sheet_name = file_trans.iloc[i, 2])
    _df['Workbook'] = file_trans.iloc[i, 1]
    _df['Sheet'] = file_trans.iloc[i, 2]
    aggregate_df1 = aggregate_df1.append(_df) 

aggregate_df1 = aggregate_df1.groupby(['TECHNOLOGY', 'FUEL', 'REGION']).sum().reset_index()

# Read in capacity data

capacity_df1 = pd.DataFrame()

# Populate the above blank dataframe with capacity data from the results workbook

for i in range(len(OSeMOSYS_filenames)):
    _df = pd.read_excel(OSeMOSYS_filenames[i], sheet_name = 'TotalCapacityAnnual')
    capacity_df1 = capacity_df1.append(_df)

# Now just extract the power capacity

pow_capacity_df1 = capacity_df1[capacity_df1['TECHNOLOGY'].str.startswith('POW')].reset_index(drop = True)

# Get maximum year column to build data frame below

year_columns = []

for item in list(aggregate_df1.columns):
    try:
        year_columns.append(int(item))
    except ValueError:
            pass

max_year = max(year_columns)

OSeMOSYS_years = list(range(2017, max_year + 1))

# Colours for charting (to be amended later)

colours = pd.read_excel('./data/2_Mapping_and_other/colour_template_7th.xlsx')
colours_hex = colours['hex']

Map_power = Map_trans[Map_trans['Sector'] == 'POW'].reset_index(drop = True)

################################ POWER SECTOR ############################### 

# Aggregate data based on the Map_power mapping

# That is group by REGION, TECHNOLOGY and FUEL

# First create empty dataframe

power_df1 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in aggregate_df1['REGION'].unique():
    interim_df1 = aggregate_df1[aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Map_power, how = 'right', on = ['TECHNOLOGY', 'FUEL'])
    interim_df1 = interim_df1.groupby(['TECHNOLOGY', 'FUEL', 'Sheet']).sum().reset_index()

    # Now add in economy reference
    interim_df1['economy'] = region
    
    # Now append economy dataframe to communal data frame 
    power_df1 = power_df1.append(interim_df1)
    
power_df1 = power_df1[['economy', 'TECHNOLOGY', 'FUEL', 'Sheet'] + OSeMOSYS_years]

Map_refownsup = Map_trans[Map_trans['Sector'].isin(['REF', 'SUP', 'OWN'])].reset_index(drop = True)

################################ REFINERY, OWN USE and SUPPLY TRANSFORMATION SECTOR ############################### 

# Aggregate data based on the Map_power mapping

# That is group by REGION, TECHNOLOGY and FUEL

# First create empty dataframe

refownsup_df1 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in aggregate_df1['REGION'].unique():
    interim_df1 = aggregate_df1[aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Map_refownsup, how = 'right', on = ['TECHNOLOGY', 'FUEL'])
    interim_df1 = interim_df1.groupby(['TECHNOLOGY', 'FUEL', 'Sheet', 'Sector']).sum().reset_index()

    # Now add in economy reference
    interim_df1['economy'] = region
    
    # Now append economy dataframe to communal data frame 
    refownsup_df1 = refownsup_df1.append(interim_df1)
    
refownsup_df1 = refownsup_df1[['economy', 'TECHNOLOGY', 'FUEL', 'Sheet', 'Sector'] + OSeMOSYS_years]

# FUEL aggregations for UseByTechnology

coal_fuel = ['1_x_coal_thermal', '1_3_lignite', '2_coal_products']
oil_fuel = ['4_5_gas_diesel_oil','4_2_naphtha', '4_6_fuel_oil', '3_1_crude_oil', '4_7_lpg', '4_8_refinery_gas_not_liq', '4_10_other_petroleum_products',]
other_fuel = ['10_electricity', '10_electricity_import','9_6_industrial_waste', '9_7_1_municipal_solid_waste_renewable', '9_7_2_municipal_solid_waste_nonrenewable', 
              '9_9_x_blackliquor', '8_1_geothermal_power', '9_4_other_biomass']
solar_fuel = ['8_2_4_solar', '8_2_1_photovoltaic']

use_agg_fuels = ['Coal', 'Oil', 'Gas', 'Hydro', 'Nuclear', 'Solar', 'Wind', 'Other']

# TECHNOLOGY aggregations for ProductionByTechnology

coal_tech = ['POW_Black_Coal_PP', 'POW_Other_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Sub_Brown_PP', 'POW_Ultra_BituCoal_PP', 'POW_CHP_COAL_PP', 'POW_Ultra_CHP_PP']
storage_tech = ['POW_AggregatedEnergy_Storage_VPP', 'POW_EmbeddedBattery_Storage']
gas_tech = ['POW_CCGT_PP', 'POW_OCGT_PP', 'POW_CHP_GAS_PP']
oil_tech = ['POW_Diesel_PP', 'POW_FuelOil_PP', 'POW_OilProducts_PP', 'POW_PetCoke_PP']
bio_tech = ['POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP']
nuclear_tech = ['POW_Nuclear_PP', 'POW_IMP_Nuclear_PP']
chp_tech = ['POW_CHP_PP']
other_tech = ['POW_Geothermal_PP', 'POW_IPP_PP', 'POW_TIDAL_PP', 'POW_WasteToEnergy_PP']
hydro_tech = ['POW_Hydro_PP', 'POW_Pumped_Hydro', 'POW_Storage_Hydro_PP', 'POW_IMP_Hydro_PP']
im_tech = ['POW_IMPORTS_PP', 'POW_IMPORT_ELEC_PP']
solar_tech = ['POW_SolarCSP_PP', 'POW_SolarFloatPV_PP', 'POW_SolarPV_PP', 'POW_SolarRoofPV_PP']
wind_tech = ['POW_WindOff_PP', 'POW_Wind_PP']

# POW_EXPORT_ELEC_PP need to work this in

prod_agg_tech = ['Coal', 'Oil', 'Gas', 'Hydro', 'Nuclear', 'Wind', 'Solar', 'Bio', 'Storage', 'Other', 'CHP', 'Imports']

# Refinery vectors

Ref_input = ['3_1_crude_oil', '3_x_NGLs']
Ref_output = ['4_1_1_motor_gasoline', '4_1_2_aviation_gasoline', '4_10_other_petroleum_products', '4_2_naphtha', '4_3_jet_fuel',
              '4_4_other_kerosene', '4_5_gas_diesel_oil', '4_6_fuel_oil', '4_7_lpg', '4_8_refinery_gas_not_liq', '4_9_ethane']

Ref_new_output = ['411_from_ref', '412_from_ref', '42_from_ref', '43_from_ref', '44_from_ref', '45_from_ref', '46_from_ref', '47_from_ref', '48_from_ref', '49_from_ref', '410_from_ref']

# Capacity vectors
    
coal_cap = ['POW_Black_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Sub_Brown_PP', 'POW_CHP_COAL_PP', 'POW_Other_Coal_PP', 'POW_Ultra_BituCoal_PP', 'POW_Ultra_CHP_PP']
gas_cap = ['POW_CCGT_PP', 'POW_OCGT_PP', 'POW_CHP_GAS_PP']
oil_cap = ['POW_Diesel_PP', 'POW_FuelOil_PP', 'POW_OilProducts_PP', 'POW_PetCoke_PP']
nuclear_cap = ['POW_Nuclear_PP', 'POW_IMP_Nuclear_PP']
hydro_cap = ['POW_Hydro_PP', 'POW_Pumped_Hydro', 'POW_Storage_Hydro_PP', 'POW_IMP_Hydro_PP']
bio_cap = ['POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP']
wind_cap = ['POW_Wind_PP', 'POW_WindOff_PP']
solar_cap = ['POW_SolarCSP_PP', 'POW_SolarFloatPV_PP', 'POW_SolarPV_PP', 'POW_SolarRoofPV_PP']
storage_cap = ['POW_AggregatedEnergy_Storage_VPP', 'POW_EmbeddedBattery_Storage']
other_cap = ['POW_CHP_PP', 'POW_Geothermal_PP', 'POW_WasteToEnergy_PP', 'POW_IPP_PP', 'POW_TIDAL_PP']
# 'POW_HEAT_HP' not in electricity capacity
transmission_cap = ['POW_Transmission']

pow_capacity_agg = ['Coal', 'Gas', 'Oil', 'Nuclear', 'Hydro', 'Biomass', 'Wind', 'Solar', 'Storage', 'Other']

# Chart years for column charts

col_chart_years = [2017, 2020, 2030, 2040, 2050]

# Define month and year to create folder for saving charts/tables

month_year = pd.to_datetime('today').strftime('%B_%Y')

# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written

# TRANSFORMATION SECTOR: Build use, capacity and production dataframes with appropriate aggregations to chart

for economy in power_df1['economy'].unique():
    use_df1 = power_df1[(power_df1['economy'] == economy) &
                        (power_df1['Sheet'] == 'UseByTechnology') &
                        (power_df1['TECHNOLOGY'] != 'POW_Transmission')].reset_index(drop = True)

    # Now build aggregate variables of the FUELS

    coal = use_df1[use_df1['FUEL'].isin(coal_fuel)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    oil = use_df1[use_df1['FUEL'].isin(oil_fuel)].groupby(['economy']).sum().assign(FUEL = 'Oil',
                                                                                    TECHNOLOGY = 'Coal power')                                                                                      

    other = use_df1[use_df1['FUEL'].isin(other_fuel)].groupby(['economy']).sum().assign(FUEL = 'Other',
                                                                                        TECHNOLOGY = 'Other power')

    solar = use_df1[use_df1['FUEL'].isin(solar_fuel)].groupby(['economy']).sum().assign(FUEL = 'Solar',
                                                                                        TECHNOLOGY = 'Solar power')

    # Use by fuel data frame 

    usefuel_df1 = use_df1.append([coal, oil, other, solar])[['FUEL',
                                                        'TECHNOLOGY'] + OSeMOSYS_years].reset_index(drop = True)

    usefuel_df1.loc[usefuel_df1['FUEL'] == '5_1_natural_gas', 'FUEL'] = 'Gas'
    usefuel_df1.loc[usefuel_df1['FUEL'] == '6_hydro', 'FUEL'] = 'Hydro'
    usefuel_df1.loc[usefuel_df1['FUEL'] == '7_nuclear', 'FUEL'] = 'Nuclear'
    usefuel_df1.loc[usefuel_df1['FUEL'] == '8_2_3_wind', 'FUEL'] = 'Wind'

    usefuel_df1 = usefuel_df1[usefuel_df1['FUEL'].isin(use_agg_fuels)].set_index('FUEL').loc[use_agg_fuels].reset_index() 

    usefuel_df1 = usefuel_df1.groupby('FUEL').sum().reset_index()
    usefuel_df1['Transformation'] = 'Input fuel'
    usefuel_df1 = usefuel_df1[['FUEL', 'Transformation'] + OSeMOSYS_years]

    nrows1 = usefuel_df1.shape[0]
    ncols1 = usefuel_df1.shape[1]

    usefuel_df2 = usefuel_df1[['FUEL', 'Transformation'] + col_chart_years]

    nrows2 = usefuel_df2.shape[0]
    ncols2 = usefuel_df2.shape[1]

    # Now build production dataframe
    prodelec_df1 = power_df1[(power_df1['economy'] == economy) &
                             (power_df1['Sheet'] == 'ProductionByTechnology') &
                             (power_df1['FUEL'].isin(['10_electricity', '10_electricity_Dx']))].reset_index(drop = True)

    # Now build the aggregations of technology (power plants)

    coal_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    oil_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(oil_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Oil')
    gas_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(gas_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas')
    storage_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(storage_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Storage')
    chp_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(chp_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'CHP')
    nuclear_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(nuclear_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(bio_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Bio')
    other_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(other_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')
    hydro_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(hydro_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Hydro')
    misc = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(im_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Imports')
    solar_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(solar_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')
    wind_pp = prodelec_df1[prodelec_df1['TECHNOLOGY'].isin(wind_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Wind')

    # Production by tech dataframe (with the above aggregations added)

    prodelec_bytech_df1 = prodelec_df1.append([coal_pp, oil_pp, gas_pp, storage_pp, chp_pp, nuclear_pp, bio_pp, other_pp, hydro_pp, misc, solar_pp, wind_pp])\
        [['TECHNOLOGY'] + OSeMOSYS_years].reset_index(drop = True)                                                                                                    

    prodelec_bytech_df1['Production'] = 'Electricity'
    prodelec_bytech_df1 = prodelec_bytech_df1[['TECHNOLOGY', 'Production'] + OSeMOSYS_years] 

    prodelec_bytech_df1 = prodelec_bytech_df1[prodelec_bytech_df1['TECHNOLOGY'].isin(prod_agg_tech)].set_index('TECHNOLOGY').loc[prod_agg_tech].reset_index()

    # CHange to TWh from Petajoules

    s = prodelec_bytech_df1.select_dtypes(include=[np.number]) / 3.6 
    prodelec_bytech_df1[s.columns] = s

    nrows3 = prodelec_bytech_df1.shape[0]
    ncols3 = prodelec_bytech_df1.shape[1]

    prodelec_bytech_df2 = prodelec_bytech_df1[['TECHNOLOGY', 'Production'] + col_chart_years]

    nrows4 = prodelec_bytech_df2.shape[0]
    ncols4 = prodelec_bytech_df2.shape[1]

    ##################################################################################################################################################################

    # Now create some refinery dataframes

    refinery_df1 = refownsup_df1[(refownsup_df1['economy'] == economy) &
                                 (refownsup_df1['Sector'] == 'REF') & 
                                 (refownsup_df1['FUEL'].isin(Ref_input))].copy()

    refinery_df1['Transformation'] = 'Input to refinery'
    refinery_df1 = refinery_df1[['FUEL', 'Transformation'] + OSeMOSYS_years]

    refinery_df1.loc[refinery_df1['FUEL'] == '3_1_crude_oil', 'FUEL'] = 'Crude oil'
    refinery_df1.loc[refinery_df1['FUEL'] == '3_x_NGLs', 'FUEL'] = 'NGLs'

    nrows5 = refinery_df1.shape[0]
    ncols5 = refinery_df1.shape[1]

    refinery_df2 = refownsup_df1[(refownsup_df1['economy'] == economy) &
                                 (refownsup_df1['Sector'] == 'REF') & 
                                 (refownsup_df1['FUEL'].isin(Ref_new_output))].copy()

    refinery_df2['Transformation'] = 'Output from refinery'
    refinery_df2 = refinery_df2[['FUEL', 'Transformation'] + OSeMOSYS_years]

    refinery_df2.loc[refinery_df2['FUEL'] == '411_from_ref', 'FUEL'] = 'Motor gasoline'
    refinery_df2.loc[refinery_df2['FUEL'] == '412_from_ref', 'FUEL'] = 'Aviation gasoline'
    refinery_df2.loc[refinery_df2['FUEL'] == '42_from_ref', 'FUEL'] = 'Naphtha'
    refinery_df2.loc[refinery_df2['FUEL'] == '43_from_ref', 'FUEL'] = 'Jet fuel'
    refinery_df2.loc[refinery_df2['FUEL'] == '44_from_ref', 'FUEL'] = 'Other kerosene'
    refinery_df2.loc[refinery_df2['FUEL'] == '45_from_ref', 'FUEL'] = 'Gas diesel oil'
    refinery_df2.loc[refinery_df2['FUEL'] == '46_from_ref', 'FUEL'] = 'Fuel oil'
    refinery_df2.loc[refinery_df2['FUEL'] == '47_from_ref', 'FUEL'] = 'LPG'
    refinery_df2.loc[refinery_df2['FUEL'] == '48_from_ref', 'FUEL'] = 'Refinery gas'
    refinery_df2.loc[refinery_df2['FUEL'] == '49_from_ref', 'FUEL'] = 'Ethane'
    refinery_df2.loc[refinery_df2['FUEL'] == '410_from_ref', 'FUEL'] = 'Other'

    refinery_df2['FUEL'] = pd.Categorical(
        refinery_df2['FUEL'], 
        categories = ['Motor gasoline', 'Aviation gasoline', 'Naphtha', 'Jet fuel', 'Other kerosene', 'Gas diesel oil', 'Fuel oil', 'LPG', 'Refinery gas', 'Ethane', 'Other'], 
        ordered = True)

    refinery_df2 = refinery_df2.sort_values('FUEL')

    nrows6 = refinery_df2.shape[0]
    ncols6 = refinery_df2.shape[1]

    refinery_df3 = refinery_df2[['FUEL', 'Transformation'] + col_chart_years]

    nrows7 = refinery_df3.shape[0]
    ncols7 = refinery_df3.shape[1]

    #####################################################################################################################################################################

    # Create some power capacity dataframes

    powcap_df1 = pow_capacity_df1[pow_capacity_df1['REGION'] == economy]

    coal_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')
    oil_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(oil_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Oil')
    wind_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(wind_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Wind')
    storage_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(storage_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Storage')
    gas_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(gas_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Gas')
    hydro_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(hydro_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Hydro')
    solar_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(solar_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Solar')
    nuclear_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(nuclear_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(bio_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Biomass')
    other_capacity = powcap_df1[powcap_df1['TECHNOLOGY'].isin(other_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Other')
    transmission = powcap_df1[powcap_df1['TECHNOLOGY'].isin(transmission_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Transmission')

    # Capacity by tech dataframe (with the above aggregations added)

    powcap_df1 = powcap_df1.append([coal_capacity, gas_capacity, oil_capacity, nuclear_capacity, hydro_capacity, bio_capacity, wind_capacity, solar_capacity, storage_capacity, other_capacity])\
        [['TECHNOLOGY'] + OSeMOSYS_years].reset_index(drop = True) 

    powcap_df1 = powcap_df1[powcap_df1['TECHNOLOGY'].isin(pow_capacity_agg)].reset_index(drop = True)

    nrows8 = powcap_df1.shape[0]
    ncols8 = powcap_df1.shape[1]

    powcap_df2 = powcap_df1[['TECHNOLOGY'] + col_chart_years]

    nrows9 = powcap_df2.shape[0]
    ncols9 = powcap_df2.shape[1]

    # Define directory
    script_dir = './results/' + month_year + '/Transformation/'
    results_dir = os.path.join(script_dir, 'economy_breakdown/', economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)

    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_transform.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None
    usefuel_df1.to_excel(writer, sheet_name = economy + '_use_fuel', index = False, startrow = chart_height)
    usefuel_df2.to_excel(writer, sheet_name = economy + '_use_fuel', index = False, startrow = chart_height + nrows1 + 3)
    prodelec_bytech_df1.to_excel(writer, sheet_name = economy + '_prodelec_bytech', index = False, startrow = chart_height)
    prodelec_bytech_df2.to_excel(writer, sheet_name = economy + '_prodelec_bytech', index = False, startrow = chart_height + nrows3 + 3)
    refinery_df1.to_excel(writer, sheet_name = economy + '_refining', index = False, startrow = chart_height)
    refinery_df2.to_excel(writer, sheet_name = economy + '_refining', index = False, startrow = chart_height + nrows5 + 3)
    refinery_df3.to_excel(writer, sheet_name = economy + '_refining', index = False, startrow = chart_height + nrows5 + nrows6 + 6)
    powcap_df1.to_excel(writer, sheet_name = economy + '_pow_capacity', index = False, startrow = chart_height)
    powcap_df2.to_excel(writer, sheet_name = economy + '_pow_capacity', index = False, startrow = chart_height + nrows8 + 3)
    
    # Access the workbook and first sheet with data from df1
    worksheet1 = writer.sheets[economy + '_use_fuel']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    worksheet1.set_column(2, ncols1 + 1, None, comma_format)
    worksheet1.set_row(chart_height, None, header_format)
    worksheet1.set_row(chart_height + nrows1 + 3, None, header_format)
    worksheet1.write(0, 0, economy + ' transformation use fuel', cell_format1)

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
                'name':       [economy + '_use_fuel', chart_height + i + 1, 0],
                'categories': [economy + '_use_fuel', chart_height, 2, chart_height, ncols1 - 1],
                'values':     [economy + '_use_fuel', chart_height + i + 1, 2, chart_height + i + 1, ncols1 - 1],
                'fill':       {'color': colours_hex[i]},
                'border':     {'none': True}
            })    
            
        worksheet1.insert_chart('B3', usefuel_chart1)

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
                'name':       [economy + '_use_fuel', chart_height + nrows1 + i + 4, 0],
                'categories': [economy + '_use_fuel', chart_height + nrows1 + 3, 2, chart_height + nrows1 + 3, ncols2 - 1],
                'values':     [economy + '_use_fuel', chart_height + nrows1 + i + 4, 2, chart_height + nrows1 + i + 4, ncols2 - 1],
                'fill':       {'color': colours_hex[i]},
                'border':     {'none': True}
            })

        worksheet1.insert_chart('J3', usefuel_chart2)

    else:
        pass

    ############################# Next sheet: Production of electricity by technology ##################################
    
    # Access the workbook and second sheet
    worksheet2 = writer.sheets[economy + '_prodelec_bytech']
    
    # Apply comma format and header format to relevant data rows
    worksheet2.set_column(2, ncols3 + 1, None, comma_format)
    worksheet2.set_row(chart_height, None, header_format)
    worksheet2.set_row(chart_height + nrows3 + 3, None, header_format)
    worksheet2.write(0, 0, economy + ' electricity production by technology', cell_format1)
    
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
                'name':       [economy + '_prodelec_bytech', chart_height + i + 1, 0],
                'categories': [economy + '_prodelec_bytech', chart_height, 2, chart_height, ncols3 - 1],
                'values':     [economy + '_prodelec_bytech', chart_height + i + 1, 2, chart_height + i + 1, ncols3 - 1],
                'fill':       {'color': colours_hex[i]},
                'border':     {'none': True}
            })    
            
        worksheet2.insert_chart('B3', prodelec_bytech_chart1)

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
                'name':       [economy + '_prodelec_bytech', chart_height + nrows3 + i + 4, 0],
                'categories': [economy + '_prodelec_bytech', chart_height + nrows3 + 3, 2, chart_height + nrows3 + 3, ncols4 - 1],
                'values':     [economy + '_prodelec_bytech', chart_height + nrows3 + i + 4, 2, chart_height + nrows3 + i + 4, ncols4 - 1],
                'fill':       {'color': colours_hex[i]},
                'border':     {'none': True}
            })    
            
        worksheet2.insert_chart('J3', prodelec_bytech_chart2)
    
    else:
        pass

    #################################################################################################################################################

    ## Refining sheet

    # Access the workbook and second sheet
    worksheet3 = writer.sheets[economy + '_refining']
    
    # Apply comma format and header format to relevant data rows
    worksheet3.set_column(2, ncols5 + 1, None, comma_format)
    worksheet3.set_row(chart_height, None, header_format)
    worksheet3.set_row(chart_height + nrows5 + 3, None, header_format)
    worksheet3.set_row(chart_height + nrows5 + nrows6 + 6, None, header_format)
    worksheet3.write(0, 0, economy + ' refining', cell_format1)

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
                'name':       [economy + '_refining', chart_height + i + 1, 0],
                'categories': [economy + '_refining', chart_height, 2, chart_height, ncols5 - 1],
                'values':     [economy + '_refining', chart_height + i + 1, 2, chart_height + i + 1, ncols5 - 1],
                'line':       {'color': colours_hex[i + 3],
                            'width': 1.25}
            })    
            
        worksheet3.insert_chart('B3', refinery_chart1)

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
                'name':       [economy + '_refining', chart_height + nrows5 + i + 4, 0],
                'categories': [economy + '_refining', chart_height + nrows5 + 3, 2, chart_height + nrows5 + 3, ncols6 - 1],
                'values':     [economy + '_refining', chart_height + nrows5 + i + 4, 2, chart_height + nrows5 + i + 4, ncols6 - 1],
                'line':       {'color': colours_hex[i],
                            'width': 1}
            })    
            
        worksheet3.insert_chart('J3', refinery_chart2)

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
                'name':       [economy + '_refining', chart_height + nrows5 + nrows6 + i + 7, 0],
                'categories': [economy + '_refining', chart_height + nrows5 + nrows6 + 6, 2, chart_height + nrows5 + nrows6 + 6, ncols7 - 1],
                'values':     [economy + '_refining', chart_height + nrows5 + nrows6 + i + 7, 2, chart_height + nrows5 + nrows6 + i + 7, ncols7 - 1],
                'fill':       {'color': colours_hex[i]},
                'border':     {'none': True}
            })    
            
        worksheet3.insert_chart('R3', refinery_chart3)

    else:
        pass

    ############################# Next sheet: Power capacity ##################################
    
    # Access the workbook and second sheet
    worksheet4 = writer.sheets[economy + '_pow_capacity']
    
    # Apply comma format and header format to relevant data rows
    worksheet4.set_column(1, ncols8 + 1, None, comma_format)
    worksheet4.set_row(chart_height, None, header_format)
    worksheet4.set_row(chart_height + nrows8 + 3, None, header_format)
    worksheet4.write(0, 0, economy + ' electricity capacity by technology', cell_format1)
    
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
                'name':       [economy + '_pow_capacity', chart_height + i + 1, 0],
                'categories': [economy + '_pow_capacity', chart_height, 1, chart_height, ncols8 - 1],
                'values':     [economy + '_pow_capacity', chart_height + i + 1, 1, chart_height + i + 1, ncols8 - 1],
                'fill':       {'color': colours_hex[i]},
                'border':     {'none': True}
            })    
            
        worksheet4.insert_chart('B3', pow_cap_chart1)

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
                'name':       [economy + '_pow_capacity', chart_height + nrows8 + i + 4, 0],
                'categories': [economy + '_pow_capacity', chart_height + nrows8 + 3, 1, chart_height + nrows8 + 3, ncols9 - 1],
                'values':     [economy + '_pow_capacity', chart_height + nrows8 + i + 4, 1, chart_height + nrows8 + i + 4, ncols9 - 1],
                'fill':       {'color': colours_hex[i]},
                'border':     {'none': True}
            })    
            
        worksheet4.insert_chart('J3', pow_cap_chart2)

    else:
        pass    

    writer.save()

print('Bling blang blaow, you have some Transformation charts now')


