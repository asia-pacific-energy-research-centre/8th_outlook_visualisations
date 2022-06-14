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

###############################################################################################

# Now grab the OSeMOSYS output emissions

# Now moving everything from OSeMOSYS to EGEDA (Only demand sectors and own use for now)

Mapping_file_emiss = Mapping_file[Mapping_file['Sector'].isin(['IND', 'OWN', 'POW', 'HYD'])].copy() 
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

#################################################################################################################

ref_industry = ref_aggemiss_df1[(ref_aggemiss_df1['TECHNOLOGY'].str.startswith('IND')) &
                                (ref_aggemiss_df1['EMISSION'].str.contains('captured'))]\
                                    .replace(np.nan, 0).reset_index(drop = True)

netz_industry = netz_aggemiss_df1[(netz_aggemiss_df1['TECHNOLOGY'].str.startswith('IND')) &
                                  (netz_aggemiss_df1['EMISSION'].str.contains('captured'))]\
                                    .replace(np.nan, 0).reset_index(drop = True)

netz_own = netz_aggemiss_df1[(netz_aggemiss_df1['TECHNOLOGY'].str.startswith('OWN')) &
                             (netz_aggemiss_df1['EMISSION'].str.contains('captured'))]\
                                .replace(np.nan, 0).reset_index(drop = True)

netz_pow = netz_aggemiss_df1[(netz_aggemiss_df1['TECHNOLOGY'].str.startswith('POW')) &
                             (netz_aggemiss_df1['TECHNOLOGY'].str.contains('CCS'))].replace(np.nan, 0)\
                                .reset_index(drop = True)

# Multiply emissions by 4 to get captured emissions (80% capture rate; 4x the amount of emissions)
netz_pow_captured = netz_pow.select_dtypes(include = [np.number]) * 4

netz_pow[netz_pow_captured.columns] = netz_pow_captured

netz_captured = netz_industry.append([netz_own, netz_pow]).reset_index(drop = True)

# Now do regional aggregations

ref_sea = ref_industry[ref_industry['REGION'].isin(['02_BD', '07_INA', '10_MAS', '15_RP', '17_SIN', '19_THA', '21_VN'])]\
    .groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()

ref_sea['REGION'] = '22_SEA'

ref_nea = ref_industry[ref_industry['REGION'].isin(['06_HKC', '08_JPN', '09_ROK', '18_CT'])]\
    .groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()

ref_nea['REGION'] = '23_NEA'

ref_oam = ref_industry[ref_industry['REGION'].isin(['03_CDA', '04_CHL', '11_MEX', '14_PE'])]\
    .groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()

ref_oam['REGION'] = '24_OAM'

ref_oce = ref_industry[ref_industry['REGION'].isin(['01_AUS', '12_NZ', '13_PNG'])]\
    .groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()

ref_oce['REGION'] = '25_OCE'

ref_captured = ref_industry.append([ref_sea, ref_nea, ref_oam, ref_oce]).reset_index(drop = True)

# Carbon neutrality

netz_sea = netz_captured[netz_captured['REGION'].isin(['02_BD', '07_INA', '10_MAS', '15_RP', '17_SIN', '19_THA', '21_VN'])]\
    .groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()

netz_sea['REGION'] = '22_SEA'

netz_nea = netz_captured[netz_captured['REGION'].isin(['06_HKC', '08_JPN', '09_ROK', '18_CT'])]\
    .groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()

netz_nea['REGION'] = '23_NEA'

netz_oam = netz_captured[netz_captured['REGION'].isin(['03_CDA', '04_CHL', '11_MEX', '14_PE'])]\
    .groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()

netz_oam['REGION'] = '24_OAM'

netz_oce = netz_captured[netz_captured['REGION'].isin(['01_AUS', '12_NZ', '13_PNG'])]\
    .groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()

netz_oce['REGION'] = '25_OCE'

netz_captured = netz_captured.append([netz_sea, netz_nea, netz_oam, netz_oce]).reset_index(drop = True)

ref_captured.to_csv(path_final + '/captured_ref.csv', index = False)
netz_captured.to_csv(path_final + '/captured_cn.csv', index = False)

print('Captured emissions dataframes saved')
