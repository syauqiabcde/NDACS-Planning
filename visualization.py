import pandas as pd
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import geopandas as gpd
import openpyxl as xl 

class Plotting:
    def __init__(self, optim_scenario):
        self.optim_scenario = optim_scenario

        self.years = np.arange(2030, 2101, 5)
        self.name_fix = {
            "United States of America": "United States",
            "Russian Federation": "Russia",
            "Viet Nam": "Vietnam",
            "Iran (Islamic Republic of)": "Iran",
            "Korea, Republic of": "South Korea",
            "Korea, Democratic People's Republic of": "North Korea",
            "Côte d'Ivoire": "Ivory Coast",
            "Dem. Rep. Congo": "DR Congo",
            "Republic of the Congo": "Congo",
            "Lao People's Democratic Republic": "Laos",
            "Syrian Arab Republic": "Syria",
            "Bolivia (Plurinational State of)": "Bolivia",
            "Venezuela (Bolivarian Republic of)": "Venezuela",
            "Tanzania, United Republic of": "Tanzania",
            'Central African Rep.': 'Central African Republic',
            'Eq. Guinea': 'Equatorial Guinea',
            'Dominican Rep.': 'Dominican Republic',
            'Czechia': 'Czech Republic',
            'N. Cyprus': 'Cyprus',
            'Trinidad and Tobago': 'Trinidad & Tobago',
            'S. Sudan': 'South Sudan',
            'Somaliland': 'Somalia',
            'Bosnia and Herz.': 'Bosnia & Herzegovina',
            'Solomon Is.': 'Solomon Islands',
            'eSwatini': 'Eswatini'
        }

        self.file_name = {0: r'Result\Result base case.xlsx',
                    1: r'Result\Result economic limitation.xlsx',
                    2: r'Result\Result technological limitation.xlsx',
                    3: r'Result\Result socio-political limitation.xlsx',
                    4: r'Result\Result resource limitation.xlsx',
                    5: r'Result\Result all limitations.xlsx'
                    }

        self.scenario_names = {0: "Base case",
                        1: "Economic limitation",
                        2: "Technological limitation",
                        3: "Socio-political limitation",
                        4: "Resource limitation",
                        5: "All limitations"
                        }

        self.map_labels = {'Capacity'    : 'DAC Capacity (GtCO₂/year)',
                'New Capacity': 'New DAC Capacity (GtCO₂/year)',
                'Nuclear Capacity': 'Nuclear Capacity (GW)',
                'New Nuclear Capacity': 'New Nuclear Capacity (GW)',
                'Investment Required': 'Investment Required (Billion USD₂₀₃₀)',
                'Export electricity': 'Export Electricity (TWh/year)',
                }

        self.plot_labels = {'CO2 ppm': 'Atmospheric CO₂ Concentration (ppm)',
                    'Investment Required': 'Investment Required (Billion USD₂₀₃₀)'}
        self.n_scenarios = len(self.scenario_names)

        # Load geographical data
        self.world = gpd.read_file(r"natural earth\ne_110m_admin_0_countries.shp")
        self.region_map = pd.read_csv("region_map.csv")
        self.world["NAME"] = self.world["NAME"].replace(self.name_fix)
        self.world = self.world.merge(
            self.region_map,
            left_on="NAME",
            right_on="country",
            how="left"
        )
        self.regions = self.world.dissolve(by="region", as_index=False)

        # Get the data
        wb = xl.load_workbook(self.file_name[self.optim_scenario]) 
        sheet_name = wb.sheetnames
        self.df = {sheet: pd.read_excel(self.file_name[self.optim_scenario], sheet_name=sheet) for sheet in sheet_name}

    def plot_map(self, parameter, year):
        results = self.df[parameter]

        results_long = results.melt(
            id_vars="Country",
            var_name="year",
            value_name=parameter
        )
        results_long.rename(columns={"Country": "region"}, inplace=True)

        results_year = results_long[
            results_long["year"] == year
        ]

        regions_plot = self.regions.merge(
            results_year,
            on="region",
            how="left"
        )

        fig, ax = plt.subplots(1, 1, figsize=(14, 8), dpi = 600)

        ax.axis("off")
        vmax = results_long[parameter].max()
        regions_plot.plot(
            column=parameter,
            ax=ax,
            cmap="YlGnBu",
            legend=False,
            edgecolor="black",
            linewidth=0.6,
            vmin= 0,
            vmax= vmax,
            missing_kwds={
                "color": "lightgrey",
                "label": "No data"
            }
        )
        sm = plt.cm.ScalarMappable(
            cmap="YlGnBu",
            norm=plt.Normalize(vmin=0, vmax=vmax)
        )
        sm._A = []

        cbar = plt.colorbar(sm, ax=ax, shrink=0.6)
        cbar.set_label(self.map_labels[parameter], fontsize=12, y=0.5)
        ax.text(0.0,1.0, f'Year: {year}', horizontalalignment='left', verticalalignment='top', transform=ax.transAxes, fontsize=12)

    def plot_line(self, parameter: str):
        plt.figure(figsize=(10,6), dpi=600)
        for optim_scenario in self.scenario_names.keys():
            wb = xl.load_workbook(self.file_name[optim_scenario]) 
            sheet_name = wb.sheetnames
            df = {sheet: pd.read_excel(self.file_name[optim_scenario], sheet_name=sheet) for sheet in sheet_name}

            results = df[parameter]
            if parameter == 'Investment Required':
                values = results.iloc[:,1:].sum(axis=0)
            else:
                values = results.iloc[:,1:].values
            
            if parameter == 'Investment Required':
                width = 0.8
                offsets = (np.arange(self.n_scenarios) - (self.n_scenarios - 1) / 2) * width
                plt.bar(self.years + offsets[optim_scenario], values, width=width, label=self.scenario_names[optim_scenario])
            else:
                plt.plot(self.years, values, label=self.scenario_names[optim_scenario], linewidth=4)
            plt.xlabel('Year')
            plt.ylabel(self.plot_labels[parameter])
            plt.legend(frameon=False)
