# NOTE: Below isn't a standalone. Reliant on things defined in Bossanova_1_actuals.py script (originally Bossanova_1.py)

# The Below is within a for loop cycling through all the APEC economies and APEC regions
# Ie builds the dataframe that gets plonked into the excel sheet (takes data from other data frames already built)

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


#########################################################################################################################################

# Then it places it an excel sheet

 # Define directory to save charts and tables workbook
    script_dir = './results/'
    results_dir = os.path.join(script_dir, economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)
        
    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_charts_' + day_month_year + '.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None

#####################################

# Here's the dataframes being placed in the sheets

    ref_kaya_1.to_excel(writer, sheet_name = 'CO2 breakdown', index = False, startrow = chart_height)
    netz_kaya_1.to_excel(writer, sheet_name = 'CO2 breakdown', index = False, startrow = chart_height + ref_kaya_1_rows + 3)

########################################################################################################################

# And then here's the chart constructions

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
            if netz_kaya_1['Carbon Neutrality'].iloc[i] in ['initial']:
                netz_kaya_chart1.add_series({
                    'name':       ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 0],
                    'categories': ['CO2 breakdown', chart_height + ref_kaya_1_rows + 3, 1, chart_height + ref_kaya_1_rows + 3, netz_kaya_1_cols - 1],
                    'values':     ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 1, chart_height + ref_kaya_1_rows + i + 4, netz_kaya_1_cols - 1],
                    'fill':       {'color': netz_kaya_1['Carbon Neutrality'].map(colours_dict).loc[i]},
                    'border':     {'none': True},
                    'gap':        50
                })

            elif netz_kaya_1['Carbon Neutrality'].iloc[i] in ['improve', 'no improve']:
                netz_kaya_chart1.add_series({
                    'name':       ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 0],
                    'categories': ['CO2 breakdown', chart_height + ref_kaya_1_rows + 3, 1, chart_height + ref_kaya_1_rows + 3, netz_kaya_1_cols - 1],
                    'values':     ['CO2 breakdown', chart_height + ref_kaya_1_rows + i + 4, 1, chart_height + ref_kaya_1_rows + i + 4, netz_kaya_1_cols - 1],
                    'fill':       {'color': netz_kaya_1['Carbon Neutrality'].map(colours_dict).loc[i],
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
