# Access the workbook and first sheet with data from df1 
    ref_worksheet28 = writer.sheets[economy + '_heat_input']
    
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
    ref_worksheet28.write(0, 0, economy + ' heat input fuel reference', cell_format1)
    ref_worksheet28.write(chart_height + ref_heat_use_2_rows + ref_heat_use_3_rows + 6, 0,\
        economy + ' heat input fuel net_zero', cell_format1)

    # Create a use by fuel area chart
    if ref_heat_use_2_rows > 0:
        heatuse_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        heatuse_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        heatuse_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        heatuse_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        heatuse_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'PJ',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        heatuse_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        heatuse_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_heat_use_2_rows):
            heatuse_chart1.add_series({
                'name':       [economy + '_heat_input', chart_height + i + 1, 0],
                'categories': [economy + '_heat_input', chart_height, 2, chart_height, ref_heat_use_2_cols - 1],
                'values':     [economy + '_heat_input', chart_height + i + 1, 2, chart_height + i + 1, ref_heat_use_2_cols - 1],
                'fill':       {'color': ref_heat_use_2['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet28.insert_chart('B3', heatuse_chart1)

    else:
        pass

    # Create a use column chart
    if ref_heat_use_3_rows > 0:
        heatuse_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        heatuse_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        heatuse_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        heatuse_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        heatuse_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'PJ',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'num_format': '# ### ### ##0',
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        heatuse_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        heatuse_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(ref_heat_use_3_rows):
            heatuse_chart2.add_series({
                'name':       [economy + '_heat_input', chart_height + ref_heat_use_2_rows + i + 4, 0],
                'categories': [economy + '_heat_input', chart_height + ref_heat_use_2_rows + 3, 2, chart_height + ref_heat_use_2_rows + 3, ref_heat_use_3_cols - 1],
                'values':     [economy + '_heat_input', chart_height + ref_heat_use_2_rows + i + 4, 2, chart_height + ref_heat_use_2_rows + i + 4, ref_heat_use_3_cols - 1],
                'fill':       {'color': ref_heat_use_3['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        ref_worksheet28.insert_chart('J3', heatuse_chart2)

    else:
        pass