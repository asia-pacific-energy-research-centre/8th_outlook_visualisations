##############################################################

    # TPES by fuel

    # access the sheet for production created above
    netz_worksheet16 = writer.sheets[economy + '_TPES_fuel_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet16.set_column(2, netz_coal_1_cols + 1, None, space_format)
    netz_worksheet16.set_row(chart_height, None, header_format)
    netz_worksheet16.set_row(chart_height + netz_coal_1_rows + 3, None, header_format)
    netz_worksheet16.set_row(chart_height + netz_coal_1_rows + netz_crude_1_rows + 6, None, header_format)
    netz_worksheet16.set_row(chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + 9, None, header_format)
    netz_worksheet16.set_row(chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + 12, None, header_format)
    netz_worksheet16.set_row(chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + 15, None, header_format)
    netz_worksheet16.set_row(chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + netz_biomass_1_rows + 18, None, header_format)
    netz_worksheet16.write(0, 0, economy + ' TPES fuel net-zero', cell_format1)
    
    # Create a TPES coal chart
    netz_tpes_coal_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_tpes_coal_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_coal_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_coal_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_coal_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Coal (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_coal_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_coal_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production', 'Imports', 'Exports', 'Stock change']:
        i = netz_coal_1[netz_coal_1['item_code_new'] == component].index[0]
        netz_tpes_coal_chart1.add_series({
            'name':       [economy + '_TPES_fuel_netz', chart_height + i + 1, 1],
            'categories': [economy + '_TPES_fuel_netz', chart_height, 2, chart_height, netz_coal_1_cols - 1],
            'values':     [economy + '_TPES_fuel_netz', chart_height + i + 1, 2, chart_height + i + 1, netz_coal_1_cols - 1],
            'fill':       {'color': netz_coal_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet16.insert_chart('B3', netz_tpes_coal_chart1)

    # Create a TPES crude oil chart
    netz_tpes_crude_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_tpes_crude_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_crude_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_crude_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_crude_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Crude oil and NGL (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_crude_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_crude_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production', 'Imports', 'Exports', 'Stock change']:
        i = netz_crude_1[netz_crude_1['item_code_new'] == component].index[0]
        netz_tpes_crude_chart1.add_series({
            'name':       [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + i + 4, 1],
            'categories': [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + 3, 2,\
                chart_height + netz_coal_1_rows + 3, netz_crude_1_cols - 1],
            'values':     [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + i + 4, 2,\
                chart_height + netz_coal_1_rows + i + 4, netz_crude_1_cols - 1],
            'fill':       {'color': netz_crude_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet16.insert_chart('J3', netz_tpes_crude_chart1)

    # Create a TPES petroleum products chart
    netz_tpes_petprod_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_tpes_petprod_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_petprod_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_petprod_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_petprod_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Petroleum products (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_petprod_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_petprod_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production', 'Imports', 'Exports', 'Bunkers', 'Stock change']:
        i = netz_petprod_2[netz_petprod_2['item_code_new'] == component].index[0]
        netz_tpes_petprod_chart1.add_series({
            'name':       [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + i + 7, 1],
            'categories': [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + 6, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + 6, netz_petprod_2_cols - 1],
            'values':     [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + i + 7, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + i + 7, netz_petprod_2_cols - 1],
            'fill':       {'color': netz_petprod_2['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet16.insert_chart('R3', netz_tpes_petprod_chart1)

    # Create a TPES gas chart
    netz_tpes_gas_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_tpes_gas_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_gas_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_gas_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_gas_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Gas (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_gas_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_gas_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production', 'Imports', 'Exports', 'Stock change']:
        i = netz_gas_1[netz_gas_1['item_code_new'] == component].index[0]
        netz_tpes_gas_chart1.add_series({
            'name':       [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + i + 10, 1],
            'categories': [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + 9, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + 9, netz_gas_1_cols - 1],
            'values':     [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + i + 10, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + i + 10, netz_gas_1_cols - 1],
            'fill':       {'color': netz_gas_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet16.insert_chart('Z3', netz_tpes_gas_chart1)

    # Create a TPES nuclear  chart
    netz_tpes_nuke_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_tpes_nuke_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_nuke_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_nuke_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_nuke_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Nuclear (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_nuke_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_nuke_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production']:
        i = netz_nuke_1[netz_nuke_1['item_code_new'] == component].index[0]
        netz_tpes_nuke_chart1.add_series({
            'name':       [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + i + 13, 1],
            'categories': [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + 12, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + 12, netz_nuke_1_cols - 1],
            'values':     [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + i + 13, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + i + 13, netz_nuke_1_cols - 1],
            'fill':       {'color': netz_nuke_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet16.insert_chart('AH3', netz_tpes_nuke_chart1)

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
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_biomass_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Biomass (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_biomass_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_biomass_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production', 'Imports', 'Exports', 'Stock change']:
        i = netz_biomass_1[netz_biomass_1['item_code_new'] == component].index[0]
        netz_tpes_biomass_chart1.add_series({
            'name':       [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + i + 16, 1],
            'categories': [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + 15, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + 15, netz_biomass_1_cols - 1],
            'values':     [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + i + 16, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + i + 16, netz_biomass_1_cols - 1],
            'fill':       {'color': netz_biomass_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet16.insert_chart('AP3', netz_tpes_biomass_chart1)

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
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_biofuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Biofuels (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'num_format': '# ### ### ##0',
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_biofuel_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_biofuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production', 'Imports', 'Exports', 'Bunkers', 'Stock change']:
        i = netz_biofuel_2[netz_biofuel_2['item_code_new'] == component].index[0]
        netz_tpes_biofuel_chart1.add_series({
            'name':       [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + netz_biomass_1_rows + i + 19, 1],
            'categories': [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + netz_biomass_1_rows + 18, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + netz_biomass_1_rows + 18, netz_biofuel_2_cols - 1],
            'values':     [economy + '_TPES_fuel_netz', chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + netz_biomass_1_rows + i + 19, 2,\
                chart_height + netz_coal_1_rows + netz_crude_1_rows + netz_petprod_2_rows + netz_gas_1_rows + netz_nuke_1_rows + netz_biomass_1_rows + i + 19, netz_biofuel_2_cols - 1],
            'fill':       {'color': netz_biofuel_2['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet16.insert_chart('AX3', netz_tpes_biofuel_chart1)