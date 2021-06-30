## Instructions for charting osemosys results (NOW UPDATED FOR 2018 BASIS YEAR)

#### 0. Before you begin

You will need several software programs on your computer:

[Visual Studio Code](https://code.visualstudio.com/) – this is a text editor that makes it easy to modify the configuration files (more on this later).

[Miniconda](https://docs.conda.io/en/latest/miniconda.html) – this is a package manager for Python. We create a specific environment (set of programs and their versions) to run the code. You want the Python 3.8 version. During installation, the installer will ask if you want to add conda to the default PATH variable. Select **YES**.

[GitHub Desktop](https://desktop.github.com/) – an easy way to grab code from GitHub. You will need to create a free account.

[Windows Terminal](https://www.microsoft.com/en-us/p/windows-terminal/9n0dx20hk701?activetab=pivot:overviewtab) – *Optional*. A modern command line terminal for Windows. You can use the built in Command Prompt too.

The following instructions assume you have installed Visual Studio Code, Miniconda, and GitHub Desktop.

#### 1. Clone ‘8th_outlook_visualisations’ from the asia-pacific-energy-research-centre Github
You can clone this repository to wherever you want on your personal computer

#### 2. In your newly cloned repository (8th_outlook_visualisations), you need to save some relevant input data in the ‘data’ folder

Save ‘EGEDA_2018_years.csv’ in 8th_outlook_visualisations/data/1_EGEDA/ 
  (available in 'Historical Energy Balances' in the 'Integration' Teams channel)

Save 
‘colours_dict.csv’

‘emissions_order_2018.csv.csv’

‘heavyind_mapping.csv’

‘macro_APEC.xlsx’

in 8th_outlook_visualisations/data/2_Mapping_and_other/
  (available in 'Charts and tables' teams folder)

Save OSeMOSYS results, (e.g. ‘07_INA_demand_results_2021-01-28-191214.xlsx’ in 8th_outlook_visualisations/data/3_OSeMOSYS_output/
  (note: 'results' needs to be lowercase)

#### 3. At the command prompt, navigate so that the ‘root’ directory is the 8th_outlook_visualisations folder
This requires the cd command to navigate to the appropriate folder executed at the command prompt (e.g. using gitbash or windows command line)

#### 4. Create the python environment once you’re in the new working directory

NB: You only need to do this once. i.e. once you've cloned the repository and created the environment with the command below, the environment is created and ready to activate (as per step 5)
```bash
conda env create --prefix ./env --file environment.yml 
```

#### 5. Activate the conda environment
```bash
conda activate ./env
```

### You're now ready to execute the scripts

#### 6. The first script to execute is in 8th_outlook_visualisations\workflow\scripts\1_historical_to_projected

NB: this needs to be run every time, as it takes the new results you've output and bolts them to EGEDA 

```bash
python ./workflow/scripts/1_historical_to_projected/OSeMOSYS_to_EGEDA_2018.py
```

Check that this has executed correctly by looking in the folder C:\Users\mathew.horne\Projects\8th_outlook_visualisations\data\4_Joined

There should be a newly created csv file 'OSeMOSYS_to_EGEDA_2018_update.csv'

#### 7. When you've run step 6 you can run the charting scripts contained in 8th_outlook_visualisations\workflow\scripts\2_charts_tables
There is currently one charts script in this folder. 

For running charts file
```bash
python ./workflow/scripts/2_charts_tables/Bossanova_1.py
```

These will create charts and tables in 8th_outlook_visualisations\results


