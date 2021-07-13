ref_refinery_1 = netz_refownsup_df1[(ref_refownsup_df1['economy'] == economy) &
                                 (ref_refownsup_df1['Sector'] == 'REF') & 
                                 (ref_refownsup_df1['FUEL'].isin(refinery_input))].copy()

    netz_refinery_1['Transformation'] = 'Input to refinery'
    netz_refinery_1 = netz_refinery_1[['FUEL', 'Transformation'] + list(ref_refinery_1.loc[:, '2017':'2050'])].reset_index(drop = True)

    netz_refinery_1.loc[netz_refinery_1['FUEL'] == 'd_ref_6_1_crude_oil', 'FUEL'] = 'Crude oil'
    netz_refinery_1.loc[netz_refinery_1['FUEL'] == 'd_ref_6_x_ngls', 'FUEL'] = 'NGLs'

    netz_refinery_1_rows = netz_refinery_1.shape[0]
    netz_refinery_1_cols = netz_refinery_1.shape[1]

    netz_refinery_2 = netz_refownsup_df1[(ref_refownsup_df1['economy'] == economy) &
                                 (ref_refownsup_df1['Sector'] == 'REF') & 
                                 (ref_refownsup_df1['FUEL'].isin(refinery_output))].copy()

    netz_refinery_2['Transformation'] = 'Output from refinery'
    netz_refinery_2 = netz_refinery_2[['FUEL', 'Transformation'] + list(ref_refinery_2.loc[:, '2017':'2050'])].reset_index(drop = True)

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
                      'Gas diesel oil', 'Fuel oil', 'LPG', 'Refinery gas', 'Ethane', 'Other'], 
        ordered = True)

    netz_refinery_2 = netz_refinery_2.sort_values('FUEL')