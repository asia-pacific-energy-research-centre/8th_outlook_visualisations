## Instructions for charting osemosys results

#### 1. Clone ‘8th_outlook_visualisations’ from the asia-pacific-energy-research-centre Github
You can clone this repository to wherever you want on your personal computer

#### 2. In your newly cloned repository (8th_outlook_visualisations), you need to save some relevant input data in the ‘data’ folder

Save ‘EGEDA_2020_June_22_wide_years_PJ.csv’ in 8th_outlook_visualisations/data/1_EGEDA/

Save ‘colour_template_7th.xlsx’ and ‘OSeMOSYS mapping.xlsx’ in 8th_outlook_visualisations/data/2_Mapping_and_other/

Save OSeMOSYS results, (e.g. ‘08_JPN_results_v1.0.xlsx’, ‘11_MEX_results_v1.2.xlsx’) in 8th_outlook_visualisations/data/3_OSeMOSYS_output/

#### 3. At the command prompt, navigate so that the ‘root’ directory is the 8th_outlook_visualisations folder
This requires the cd command to navigate to the appropriate folder executed at the command prompt (e.g. using gitbash)

#### 4. Create the python environment once you’re in the new working directory
```bash
conda env create --prefix ./env --file ./workflow/environment.yml 
```

#### 5. Activate the conda environment
```bash
source activate ./env
```

### You're now ready to execute the scripts

#### 6. The first script to execute is in 8th_outlook_visualisations\workflow\scripts\1_historical_to_projected

NB: this needs to be run every time, as it takes the new results you've output and bolts them to EGEDA 

```bash
python ./workflow/scripts/1_historical_to_projected/OSeMOSYS_to_EGEDA.py
```

Check that this has executed correctly by looking in the folder C:\Users\mathew.horne\Projects\8th_outlook_visualisations\data\4_Joined

There should be a newly created csv file 'OSeMOSYS_to_EGEDA.csv'

#### 7. When you've run step 6 you can run the charting scripts contained in 8th_outlook_visualisations\workflow\scripts\2_charts_tables
There are currently four scripts in this folder. 

Example for running TPES
```bash
python ./workflow/scripts/2_charts_tables/TPES_economy.py
```
This will create TPES charts and tables in 8th_outlook_visualisations\results


