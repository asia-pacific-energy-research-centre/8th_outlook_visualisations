## Instructions for charting osemosys results (NOW UPDATED FOR 2018 BASIS YEAR)

### 0. Before you begin

You will need several software programs on your computer:

[Visual Studio Code](https://code.visualstudio.com/) – this is a text editor that makes it easy to modify the configuration files (more on this later).

[Miniconda](https://docs.conda.io/en/latest/miniconda.html) – this is a package manager for Python. We create a specific environment (set of programs and their versions) to run the code. You want the Python 3.8 version. During installation, the installer will ask if you want to add conda to the default PATH variable. Select **YES**.

[GitHub Desktop](https://desktop.github.com/) – an easy way to grab code from GitHub. You will need to create a free account.

[Windows Terminal](https://www.microsoft.com/en-us/p/windows-terminal/9n0dx20hk701?activetab=pivot:overviewtab) – *Optional*. A modern command line terminal for Windows. You can use the built in Command Prompt too.

The following instructions assume you have installed Visual Studio Code, Miniconda, and GitHub Desktop.

### 1. Clone ‘8th_outlook_visualisations’
Create a folder on your computer called `GitHub`. 

Open your command prompt and navigate to that folder, for example: `cd GitHub`. 

Clone this repository by typing `git clone https://github.com/asia-pacific-energy-research-centre/8th_outlook_visualisations.git`

### 2. Install the necessary software in the 8th_outlook_visualisations folder

Navigate into the folder by typing `cd 8th_outlook_visualisations`

Create the environment using `conda env create --prefix ./env --file ./environment.yml`

### 3. Add the necessary files

### 2020 UPDATE: YOU WILL NEED TO SAVE UPDATED EGEDA 2020 DATA AND NEW MAPPING FILE 

Copy the following files from the Integration\Historical Energy Balances\2018 folder:
- EGEDA_2018_years
- EGEDA_FC_CO2_Emissions_years_2018
- NEW: EGEDA_2020_created_21102022

(NEW EGEDA data set is from the Outlook 9h Modelling Teams channel)

Save those files in 8th_outlook_visualisations\data\1_EGEDA.

Copy the following files from the Charts and tables\Mapping:
- colours_dict
- heavyind_mapping
- macro_APEC
- OSeMOSYS_mapping_2021
- NEW: OSeMOSYS_mapping_for_2020_data

Save those files in 8th_outlook_visualisations\data\2_Mapping_and_other.

Take your results files and save them in 8th_outlook_visualisations\data\3_OSeMOSYS_output.

### 4. Run the script

### UPDATE: THERE ARE NEW SCRIPTS THAT INCORPORATE 2020 EGEDA data

Ensure that the environment is activated by executing at the command line: 'soure activate ./env' (you should already have changed the directory to '8th_outlook_visualisations'

For preparing relevant files that are used for charting, this script is common (run command):
python ./workflow/scripts/1_historical_to_projected/OSeMOSYS_to_EGEDA_2018_actuals.py

Then to run the most sought after charts script:
python ./workflow/scripts/2_charts_tables/Bossanova_1_actuals.py

Note, there are other scripts available in the workflow folder (some are work in progress).

THE 2020 EGEDA DATA SCRIPTS ARE:

python ./workflow/scripts/1_historical_to_projected/OSeMOSYS_to_EGEDA_2020_actuals.py
python ./workflow/scripts/2_charts_tables/Bossanova_1_with_2020_historical.py

### 5. View charts

Charts are saved in the 8th_outlook_visualisations\results folder.
