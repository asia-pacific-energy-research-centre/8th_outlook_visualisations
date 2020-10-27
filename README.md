## Instructions for charting osemosys results

##### NB: see generic readme (below these seven steps) for tips on working with github repositories

#### 1. Clone ‘8th_outlook_visualisations’ from the asia-pacific-energy-research-centre Github
You can clone this repository to wherever you want on your personal computer

#### 2. In your newly cloned repository (8th_outlook_visualisations), you need to save some relevant input data in the ‘data’ folder

Save ‘EGEDA_2020_June_22_wide_years_PJ.csv’ in 8th_outlook_visualisations/data/1_EGEDA/
Save ‘colour_template_7th.xlsx’ and ‘OSeMOSYS mapping.xlsx’ in 8th_outlook_visualisations/data/2_Mapping_and_other/
Save OSeMOSYS results, (e.g. ‘08_JPN_results_v1.0.xlsx’, ‘11_MEX_results_v1.2.xlsx’) in 8th_outlook_visualisations/data/3_OSeMOSYS_output/

#### 3. At the command prompt, navigate so that the ‘root’ directory is the 8th_outlook_visualisations folder
This requires the cd command to navigate to the appropriate folder executed at the command prompt (e.g. using gitbash)

#### 4. Create the python environment once you’re in the new working directory
conda env create --prefix ./env --file ./workflow/environment.yml 

#### 5. Activate the conda environment
source activate ./env   

### You're now ready to execute the scripts

#### 6. The first script to execute is in 8th_outlook_visualisations\workflow\scripts\1_historical_to_projected
It is called OSeMOSYS_to_EGEDA.py and will take the results files you saved in step 2 above and bolt them to the EGEDA historical data
Execute at command prompt:
python ./workflow/scripts/1_historical_to_projected/OSeMOSYS_to_EGEDA.py

Check that this has executed correctly by looking in the folder C:\Users\mathew.horne\Projects\8th_outlook_visualisations\data\4_Joined
There should be a newly created csv file 'OSeMOSYS_to_EGEDA.csv'

#### 7. When you've run step 6 you can run the charting scripts contained in 8th_outlook_visualisations\workflow\scripts\2_charts_tables

There are currently four scripts in this folder. 
Run them just like above. For example, to run TPES, at the command prompt, execute: python ./workflow/scripts/2_charts_tables/TPES_economy.py

This will create TPES charts and tables in this folder: 8th_outlook_visualisations\results

### Template details

## aperc-template
Template for APERC models.

### How to use this template
Create a new repository. When given the option, select 'aperc-template' as the template.

### Project organization

Project organization is based on ideas from [_Good Enough Practices for Scientific Computing_](https://journals.plos.org/ploscompbiol/article?id=10.1371/journal.pcbi.1005510) and the [_SnakeMake_](https://snakemake.readthedocs.io/en/stable/snakefiles/deployment.html) recommended workflow. 

1. Put each project in its own directory, which is named after the project.
2. Put data in the `data` directory. This can be input data or data files created by scripts and notebooks in this project.
3. Put configuration files in the `config` directory.
4. Put text documents associated with the project in the `docs` directory.
5. Put all scripts in the `workflow/scripts` directory.
6. Install the Conda environment into the `workflow/envs` directory. 
7. Put all notebooks in the `workflow/notebooks` directory.
8. Put final results in the `results` directory.
9. Name all files to reflect their content or function.

### Using Conda

#### Creating the Conda environment

After adding any necessary dependencies to the Conda `environment.yml` file you can create the 
environment in a sub-directory of your project directory by running the following command.

```bash
$ conda env create --prefix ./env --file ./workflow/environment.yml
```
Once the new environment has been created you can activate the environment with the following 
command.

```bash
$ conda activate ./env
```

Note that the `env` directory is *not* under version control as it can always be re-created from 
the `environment.yml` file as necessary.

#### Updating the Conda environment

If you add (remove) dependencies to (from) the `environment.yml` file after the environment has 
already been created, then you can update the environment with the following command.

```bash
$ conda env update --prefix ./env --file ./workflow/environment.yml --prune
```

#### Listing the full contents of the Conda environment

The list of explicit dependencies for the project are listed in the `environment.yml` file. To see the full list of packages installed into the environment run the following command.

```bash
conda list --prefix ./env
```

