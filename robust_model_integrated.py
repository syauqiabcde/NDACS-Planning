#%%
import numpy as np
from pyomo.util.infeasible import log_infeasible_constraints
from pyomo.environ import *
import pandas as pd
import openpyxl as xl 
import time
from scipy.interpolate import RegularGridInterpolator
from amplpy import modules

scenario_dict = {0: 'base case',
                 1: 'economic limitation',
                 2: 'technological limitation',
                 3: 'socio-political limitation',
                 4: 'resource limitation',
                 5: 'all limitations'}

def run_model(co2_scenario='IPCC middle', ppm_limit_scenario='1.5 C', optim_scenario=None):
    print("Running optimization model")
    start = time.time()

    # Case name
    case_name = scenario_dict[optim_scenario]
    print(f"case name: {case_name} - {co2_scenario}")

    # Scenario
    scenario = co2_scenario 
    ppm_limit ={'1.5 C': (445.0, 390.0),      # IPCC, 2018, Global warming of 1.5 °C, Intergovernmental Panel on Climate Change, Geneva.
            '2 C': (470.0, 460.0)}

    # General data, asumptions, and constants
    ppm_to_mass = 2.13      # Conversion from ppm to mass CO2, 1 ppm of CO2 = 2.13 giga ton of CO2
    percent_allowable_investment = 0.01
    percent_allowable_om = 0.005
    plant_lifetime = 30
    nuclear_power_plant_lifetime = 60
    water_consumption = 1 # in billion m3/Gton CO2/year              10.1016/j.rser.2023.113215
    nuclear_water_consumption = 8.9e-3         # in billion m3/GW          https://docs.nrel.gov/docs/fy11osti/50900.pdf
    land_consumption = 0.2*1e9/1e6 # in km2/Gt CO2/year      10.1016/j.cej.2024.154421
    nuclear_land_consumption = 0.34 # in km2/GW      https://doi.org/10.1371/journal.pone.0270155
    uranium_reserve = 8e6  # in ton
    uranium_consumption = 6.89343e-7 # ton/GJ
    
    discount_rate = 0.1 # 10% of discount rate
    fixed_om = 0.037 # 3.7% fixed O&M cost of Nuclear DACCS plant
    nuclear_fixed_om = 0.05 # 5% fixed O&M cost of Nuclear power plant
    cf_maximum = 0.95 # maximum capacity factor due to maintenance and unplanned shutdown
    cf_minimum = 0.1 # minimum capacity factor due to technical limitation
    H = 24
    nuclear_emission_factor = 12 / 0.0036 / 1e6 # in ton/GJ, 12 gCO2/kWh, 0.0036e6 conversion from kWh to GJ
    elc_cost = 0.0154 # $/MJ
    heat_cost = 0.0067 # $/MJ
    heat_share = 0.75   # percentage of heat energy requirement compared to total energy requirement / elc_share = 1-heat_share
    co2_peak_year = 2050

    # Getting data from excel
    file_name = 'data_5years.xlsx'
    wb = xl.load_workbook(file_name)
    sheet_name = wb.sheetnames
    df = {sheet: pd.read_excel(file_name, sheet_name=sheet) for sheet in sheet_name}
    years = df['capex'].iloc[:,0].values
    countries = df['env_condition'].iloc[:,0].values
    ccs_included = df['ccs_included'].iloc[:,:].values.reshape(-1)
    non_NPT = df['non-NPT'].iloc[:,:].values.reshape(-1)
    duty = df['energy_duty_adjusted'].iloc[:,1:].values # in GJ/ton CO2
    temperature = np.linspace(0, 40, 17)
    rh = np.linspace(0, 100, 21)
    co2_tax = df['co2_tax'].iloc[:,1:].values           # in $/ton CO2
    co2_storage_capacity = df['co2_storage_capacity'].iloc[:,1].values  # in Gton CO2   
    gdp = df['gdp'].iloc[:,1:].values   # in million USD
    plant_capex = df['capex'].iloc[:,1].values  # in $/ton CO2/year
    ccs_readiness_index = df['ccs_readiness'].iloc[:,1].values
    nuclear_readiness_index = df['nuclear_readiness'].iloc[:,1].values
    nuclear_support = df['nuclear_perception'].iloc[:,1].values
    nuclear_oppose = df['nuclear_perception'].iloc[:,2].values
    water_reserve = df['resource_reserve'].iloc[:,1].values     # in billion m3/year
    land_reserve = df['resource_reserve'].iloc[:,2].values      # in km2
    nuclear_capex = df['nuclear_capex'].iloc[:,1].values
    emission_factor = df['electricity_emission_factor'].iloc[:,1].values    # in ton CO2/GJ
    electricity_consumed = df['electricity_consumed'].iloc[:,1:].values     # in GJ
    length = df['interval'].iloc[:,1].values
    interval = {years[i]: length[i] for i in range(len(years))}
    uranium_price = df['uranium_price'].iloc[:,1].values  # in $/ton

    # Weather data
    weather_years = [2011, 2013, 2014, 2015, 2017, 2018, 2019, 2021, 2022, 2023]

    num_hours = pd.read_excel('weather_data.xlsx', sheet_name=f"temperature_data_{weather_years[0]}").shape[1] - 1  
    num_periods = len(years)
    num_countries = len(countries)

    # Get maximum duty for the worst case scenario
    interpolator = RegularGridInterpolator((rh,temperature), duty, bounds_error=False, fill_value=None)

    energy_duty_data = []

    for i, country in enumerate(countries):
        weather_data = []
        for year in weather_years:
            # Load temperature and humidity for the year
            temp_df = pd.read_excel('weather_data.xlsx', sheet_name=f"temperature_data_{year}")
            rh_df   = pd.read_excel('weather_data.xlsx', sheet_name=f"humidity_data_{year}")
            
            temp_row = temp_df.iloc[i, 1:].values 
            rh_row   = rh_df.iloc[i, 1:].values
            weather_row = np.stack((temp_row, rh_row))
            weather_data.append(weather_row)

        weather_array = np.stack(weather_data, axis=2)

        T  = weather_array[0] 
        RH = weather_array[1]  

        points = np.column_stack([RH.ravel(), T.ravel()])  
        duty_all = interpolator(points)   
        duty_all = duty_all.reshape(num_hours, -1)                
        duty_max_per_hour = duty_all.max(axis=1)
        energy_duty_data.append(duty_max_per_hour)

    duty_data_array = np.stack(energy_duty_data) 
    max_duty = np.max(duty_data_array, axis=1)  # in MJ/kg CO2
    power_plant_efficiency = 0.33
    power_plant_cf  = 0.92

    duty_data = {(countries[i], h): duty_data_array[i, h] for i in range(num_countries) for h in range(num_hours)}
    max_duty_data = {countries[i]: max_duty[i] for i in range(num_countries)}
    co2_tax_data = {(countries[i], years[t]): co2_tax[i, t] for i in range(num_countries) for t in range(num_periods)}
    co2_storage_data = {countries[i]: co2_storage_capacity[i] for i in range(num_countries)}
    gdp_data = {(countries[i], years[t]): gdp[i, t] for i in range(num_countries) for t in range(num_periods)}
    capex_data = {years[t]: plant_capex[t] for t in range(num_periods)}
    nuclear_capex_data = {years[t]: nuclear_capex[t] for t in range(num_periods)}
    ccs_readiness_data = {countries[i]: ccs_readiness_index[i] for i in range(num_countries)}
    nuclear_readiness_data = {countries[i]: nuclear_readiness_index[i] for i in range(num_countries)}
    nuclear_support_data = {countries[i]: nuclear_support[i] for i in range(num_countries)}
    nuclear_oppose_data = {countries[i]: nuclear_oppose[i] for i in range(num_countries)}
    water_reserve_data = {countries[i]: water_reserve[i] for i in range(num_countries)}
    land_reserve_data = {countries[i]: land_reserve[i] for i in range(num_countries)}
    emission_factor_data = {countries[i]: emission_factor[i] for i in range(num_countries)}
    electricity_consumed_data = {(countries[i], years[t]): electricity_consumed[i, t] for i in range(num_countries) for t in range(num_periods)}
    uranium_price_data = {years[t]: uranium_price[t] for t in range(num_periods)}

    # Define model
    model = ConcreteModel()

    # Sets
    model.i = Set(initialize=countries)
    model.i_ccs_included = Set(initialize=ccs_included)
    model.i_non_NPT = Set(initialize=non_NPT)
    model.t = Set(initialize=years)
    model.h = Set(initialize=range(0, num_hours))

    # Parameters
    model.energy_duty = Param(model.i, model.h, initialize=duty_data, within=Reals)
    model.max_duty = Param(model.i, initialize=max_duty_data, within=Reals)
    model.co2_tax = Param(model.i, model.t, initialize=co2_tax_data, within=Reals)
    model.co2_storage = Param(model.i, initialize=co2_storage_data , within=Reals)
    model.gdp = Param(model.i, model.t, initialize=gdp_data , within=Reals)
    model.capex = Param(model.t, initialize=capex_data , within=Reals)
    model.nuclear_capex = Param(model.t, initialize=nuclear_capex_data, within=Reals)
    model.ccs_readiness = Param(model.i, initialize=ccs_readiness_data, within=Reals)
    model.nuclear_readiness = Param(model.i, initialize=nuclear_readiness_data, within=Integers)
    model.nuclear_support = Param(model.i, initialize=nuclear_support_data, within=Reals)
    model.nuclear_oppose = Param(model.i, initialize=nuclear_oppose_data, within=Reals)    
    model.water_reserve = Param(model.i, initialize=water_reserve_data, within=Reals)
    model.land_reserve = Param(model.i, initialize=land_reserve_data, within=Reals)
    model.emission_factor = Param(model.i, initialize=emission_factor_data, within=Reals)
    model.electricity_consumed = Param(model.i, model.t, initialize=electricity_consumed_data, within=Reals)
    model.uranium_price = Param(model.t, initialize=uranium_price_data, within=Reals)

    print("Parameters initialization finished.")

    #%% Model construction

    # Decision Variables
    model.new = Var(model.i, model.t, within=NonNegativeReals, bounds=(0.0, 0.3))   # new DACCS plant capacity in Gton CO2/year
    model.co2_ppm = Var(model.t, within=NonNegativeReals)       # not really a decision variables just a way to determine the co2 ppm based on the previous value of co2 ppm
    model.u = Var(model.i, model.t, within=NonNegativeReals)               # utilization of DACCS plant in Gton CO2/year, result of cf * capacity (used for avoiding bilinear term in objective function)
    model.new_nuclear = Var(model.i, model.t, within=NonNegativeReals, bounds=(0.0, 150.0)) # new nuclear power plant capacity in GW
    model.nuclear_req = Var(model.i, model.t, within=NonNegativeReals)                  # nuclear power plant capacity required in GW
    model.nuclear_shortfall = Var(model.i, model.t, within=NonNegativeReals)            # nuclear power plant capacity shortfall in GW
    model.export_elc = Var(model.i, model.t, model.h, within=NonNegativeReals)          # excess electricity exported in GJ

    # Expressions
    def IPCC_prediction(t, scenario='IPCC middle'):        # return the CO2 concentration in ppm based on IPCC prediction 
        if scenario == 'IPCC conservative':
            return 0.03125 * t**2 - 121.37500 * t + 118027.5
        elif scenario == 'IPCC middle':
            return 0.00537 * t**2 - 19.36899 * t + 17614.35624
        elif scenario == 'IPCC optimistic':
            return -0.01563 * t**2 + 66.31250 * t - 69801.25000
        
    def retirement(model, i, t):
        retirement_period = t - plant_lifetime  # Find the period that corresponds to retirement period
        if retirement_period in model.t:
            return model.new[i, retirement_period]  # Retire what has been exceed lifetime
        else:
            return 0
            
    def capacity_expansion(model, i, t):
        if t == min(model.t):
            return model.new[i,t]
        else:
            return model.capacity[i,t-interval[t]] + model.new[i,t] - model.retired[i,t]

    def captured_co2(model, i, t, h):   # hourly CO2 captured
        return model.u[i,t,h] / num_hours * interval[t]
        
    def reduced_co2(model, i, t, h):
        return (model.emission_factor[i] - nuclear_emission_factor) * model.export_elc[i,t,h] * interval[t] / 1e9  # in giga ton CO2
        
    def nuclear_elc(model, i,t):
        return model.capacity[i,t] * model.max_duty[i] * 1e12 * (1 - heat_share) / (power_plant_cf * 8760 * 3600 * 1000)    # Power plant capacity need based on electricity demand from DACCS, 8760*3600*1000 conversion from MJ/year to GW

    def nuclear_heat(model, i,t):
        return model.capacity[i,t] * model.max_duty[i] * 1e12 * heat_share * (power_plant_efficiency / (1-power_plant_efficiency)) / (power_plant_cf * 8760 * 3600 * 1000)    # Power plant capacity need based on heat demand from DACCS, 8760*3600*1000 conversion from MJ/year to GW   

    def nuclear_retirement(model, i, t):
        retirement_period = t - nuclear_power_plant_lifetime  # Find the period that corresponds to retirement period
        if retirement_period in model.t:
            return model.new_nuclear[i, retirement_period]  # Retire what has been exceed lifetime
        else:
            return 0

    def nuclear_capacity_rule(model, i, t):
        if t == min(model.t):
            return model.new_nuclear[i,t]
        else:
            return model.nuclear_capacity[i,t-interval[t]] + model.new_nuclear[i,t] - model.nuclear_retired[i,t]

    def excess_electricity(model, i, t, h):
        return model.electricity_production[i,t,h] - model.electricity_consumption[i,t,h]
        
    def system_om_cost(model, i, t):
        daccs_fix_om = model.capex[t] * 1e9 * model.capacity[i,t] * fixed_om           # Fixed OM cost of Nuclear DACCS plant
        nuclear_fix_om = model.nuclear_capex[t] * 1e6 * model.nuclear_capacity[i,t] * nuclear_fixed_om                     # Nuclear power plant investment cost
        uranium_cost = sum(model.uranium_pp_consumption[i,t,h]  * model.uranium_price[t] for h in model.h)
        return (daccs_fix_om + nuclear_fix_om + uranium_cost) * interval[t]

    def system_land_consumption(model, i, t):
        return model.capacity[i,t] * land_consumption + model.nuclear_capacity[i,t] * nuclear_land_consumption

    def system_water_consumption(model, i, t):
        return model.capacity[i,t] * water_consumption + model.nuclear_capacity[i,t] * nuclear_water_consumption

    def electricity_production(model, i, t, h):
        return model.nuclear_capacity[i,t] * power_plant_cf * 8760 * 3600 / num_hours
    
    def electricity_consumption(model, i, t, h):
        return model.energy_duty[i,h] * (1-heat_share) * 1e12 * model.captured_co2[i,t,h]  / 1000  # electricity consumed by DACCS plant, 1000 conversion from MJ to GJ   

    def uranium_pp_consumption(model, i, t, h):
        return model.electricity_production[i,t,h] * uranium_consumption

    model.retired = Expression(model.i, model.t, rule=retirement)
    model.capacity = Expression(model.i, model.t, rule=capacity_expansion)
    model.captured_co2 = Expression(model.i, model.t, model.h, rule=captured_co2)
    model.reduced_co2 = Expression(model.i, model.t, model.h, rule=reduced_co2)
    model.nuclear_elc = Expression(model.i, model.t, rule=nuclear_elc)
    model.nuclear_heat = Expression(model.i, model.t, rule=nuclear_heat)
    model.nuclear_retired = Expression(model.i, model.t, rule=nuclear_retirement)
    model.nuclear_capacity = Expression(model.i, model.t, rule=nuclear_capacity_rule)
    model.electricity_production = Expression(model.i, model.t, model.h, rule=electricity_production)
    model.electricity_consumption = Expression(model.i, model.t, model.h, rule=electricity_consumption)
    model.uranium_pp_consumption = Expression(model.i, model.t, model.h, rule=uranium_pp_consumption)
    model.excess_electricity = Expression(model.i, model.t, model.h, rule=excess_electricity)
    model.om_cost = Expression(model.i, model.t, rule=system_om_cost)
    model.land_consumption = Expression(model.i, model.t, rule=system_land_consumption)
    model.water_consumption = Expression(model.i, model.t, rule=system_water_consumption)
    
    # Objective Function
    def tot_cost(model):
        return sum(((model.capex[t] * 1e9 * model.new[i,t]                                      # DACCS investment cost
                    + model.nuclear_capex[t] * 1e6 * model.new_nuclear[i,t]                     # Nuclear power plant investment cost
                    + model.om_cost[i,t])                                                     # Total operation and maintenance cost of the system
                    - model.co2_tax[i, t] * 1e9 * sum(model.captured_co2[i,t,h] + model.reduced_co2[i,t,h] for h in model.h)                     # CO2 tax benefit
                    - sum(model.export_elc[i,t,h] * elc_cost * 1e3 for h in model.h))                      # benefit from selling excess electricity
                    / (1+discount_rate)**(t-min(model.t))                                       # Discount rate  
                    for i in model.i for t in model.t) / 1e12
            
    model.obj = Objective(rule=tot_cost, sense=minimize)

    # Mandatory Constraints
    def co2_ppm_limit(model, t):
        if t == max(model.t):
            return model.co2_ppm[t] <= ppm_limit[ppm_limit_scenario][1]        # the co2 ppm in 2100 must lower than 411 ppm (Paris agreement 1.5 C scenario) 
        else:
            return model.co2_ppm[t] <= ppm_limit[ppm_limit_scenario ][0]        # the co2 ppm in all years must lower than 465 ppm (Paris agreement 1.5 C scenario) 

    def co2_ppm_trend(model, t):
        if t < co2_peak_year:
            return Constraint.Skip
        else:
            return model.co2_ppm[t] <= model.co2_ppm[t-interval[t]]
    
    def co2_storage(model, i):
        return sum(model.captured_co2[i,t,h] for t in model.t for h in model.h) <= model.co2_storage[i]

    def co2_ppm_rule(model, t):
        if t == min(model.t):  # Initial CO2 concentration at the first period
            captured = sum(model.captured_co2[i, t, h] for i in model.i for h in model.h) 
            reduced = sum(model.reduced_co2[i, t, h] for i in model.i for h in model.h) # in giga ton CO2
            emitted = sum(model.electricity_production[i,t,h] * nuclear_emission_factor * interval[t] / 1e9 for i in model.i for h in model.h)  # in giga ton CO2)
            return model.co2_ppm[t] == (IPCC_prediction(t, scenario) * ppm_to_mass + emitted - captured - reduced) / ppm_to_mass
        else:
            generated = (IPCC_prediction(t, scenario) - IPCC_prediction(t-interval[t]+1, scenario)) * ppm_to_mass
            captured = sum(model.captured_co2[i, t, h] for i in model.i for h in model.h) 
            reduced = sum(model.reduced_co2[i, t, h] for i in model.i for h in model.h) # in giga ton CO2
            emitted = sum(model.electricity_production[i,t,h] * nuclear_emission_factor * interval[t] / 1e9 for i in model.i for h in model.h)  # in giga ton CO2)
            return model.co2_ppm[t] == model.co2_ppm[t-interval[t]] + (generated + emitted - captured - reduced) / ppm_to_mass

    def rampup_constraint(model, i, t, h):
        if h % H == 0:
            return Constraint.Skip
        else:
            return (model.captured_co2[i, t, h] - model.captured_co2[i, t, h-1]) / interval[t] <= 0.1 * model.capacity[i, t]
        
    def rampdown_constraint(model, i, t, h):
        if h % H == 0:
            return Constraint.Skip
        else:
            return (model.captured_co2[i, t, h-1] - model.captured_co2[i, t, h]) / interval[t] <= 0.1 * model.capacity[i, t]

    def upper_u(model, i, t, h):
        return model.captured_co2[i, t, h] / interval[t] <= model.capacity[i, t] * cf_maximum

    def lower_u(model, i, t, h):
        return model.captured_co2[i, t, h] / interval[t] >= model.capacity[i, t] * cf_minimum

    def nuclear_capacity_elc(model, i, t):
        return model.nuclear_req[i,t] >= model.nuclear_elc[i,t]

    def nuclear_capacity_heat(model, i, t):
        return model.nuclear_req[i,t] >= model.nuclear_heat[i,t]

    def just_enough_nuclear(model, i, t):
        if t == min(model.t):
            prev = 0
        else:
            prev = model.nuclear_capacity[i, t-interval[t]]
        return model.nuclear_shortfall[i,t] >= model.nuclear_req[i,t] - prev + model.nuclear_retired[i,t]

    def new_nuclear_equals_shortfall(model, i, t):
        return model.new_nuclear[i,t] == model.nuclear_shortfall[i,t]

    def export_upper_limit(model, i, t, h):
        return model.export_elc[i,t,h] <= model.excess_electricity[i,t,h]

    def export_consumption_limit(model, i, t):
        return sum(model.export_elc[i,t,h] for h in model.h) <= model.electricity_consumed[i,t]

    def uranium_resource(model):
        return sum(model.uranium_pp_consumption[i,t,h] for i in model.i for t in model.t for h in model.h) <= uranium_reserve
    
    model.ppm_limit = Constraint(model.t, rule=co2_ppm_limit)
    model.co2_ppm_trend_constraint = Constraint(model.t, rule=co2_ppm_trend)
    model.storage_constraint = Constraint(model.i, rule=co2_storage)
    model.co2_atm_balance = Constraint(model.t, rule=co2_ppm_rule)
    model.rampup_constrant = Constraint(model.i, model.t, model.h, rule=rampup_constraint)
    model.rampdown_constrant = Constraint(model.i, model.t, model.h, rule=rampdown_constraint)
    model.upper_u_constraint = Constraint(model.i, model.t, model.h, rule=upper_u)
    model.loweru_constraint = Constraint(model.i, model.t, model.h, rule=lower_u)
    model.nuclear_capacity_elc_constraint = Constraint(model.i, model.t, rule=nuclear_capacity_elc)
    model.nuclear_capacity_heat_constraint = Constraint(model.i, model.t, rule=nuclear_capacity_heat)
    model.just_enough_nuclear_elc_constraint = Constraint(model.i, model.t, rule = just_enough_nuclear)
    model.new_nuclear_equals_shortfall_constraint = Constraint(model.i, model.t, rule = new_nuclear_equals_shortfall)
    model.export_elc_excess_constraint = Constraint(model.i, model.t, model.h, rule=export_upper_limit)
    model.export_elc_consumption_constraint = Constraint(model.i, model.t, rule=export_consumption_limit)
    model.uranium_resource_constraint = Constraint(rule=uranium_resource)

    # Optional Constraints
    def financing_constraint(model, i, t):
        return (model.new[i, t] * model.capex[t] * 1e9 + model.nuclear_capex[t] * 1e6 * model.new_nuclear[i,t]) <= percent_allowable_investment * model.gdp[i, t] * 1e6
    
    def opex_financing_constraint(model, i, t):
        return model.om_cost[i,t] <= percent_allowable_om * model.gdp[i, t] * 1e6
    
    def ccs_readiness_constraint(model, i, t):
        if model.ccs_readiness[i] < 50:
            return model.new[i,t] == 0
        return Constraint.Skip

    def nuclear_readiness_constraint(model, i, t):
        if model.nuclear_readiness[i] < 3:
            return model.new[i,t] == 0
        return Constraint.Skip

    def countries_NPT(model, i, t):
        if i in model.i_non_NPT.data():
            return model.new[i,t] == 0
        return Constraint.Skip

    def public_perception_constraint(model, i, t):
        if model.nuclear_support[i] < model.nuclear_oppose[i]:
            return model.new[i,t] == 0
        return Constraint.Skip
    
    def ccs_included_constraint(model, i, t):
        if i not in model.i_ccs_included:
            return model.new[i,t] == 0
        return Constraint.Skip

    def water_reserve_constraint(model, i, t):
        return model.water_consumption[i,t] <= model.water_reserve[i]

    def land_reserve_constraint(model, i, t):
        return model.land_consumption[i,t] <= model.land_reserve[i]

    # Economic aspect
    if optim_scenario == 1 or optim_scenario == 5:
        model.financing_constraint = Constraint(model.i, model.t, rule=financing_constraint)
        model.opex_financing_constraint = Constraint(model.i, model.t, rule=opex_financing_constraint)

    # Technological readiness
    if optim_scenario == 2 or optim_scenario == 5:
        model.ccs_readiness_constraint = Constraint(model.i, model.t, rule=ccs_readiness_constraint)
        model.nuclear_readiness_constraint = Constraint(model.i, model.t, rule=nuclear_readiness_constraint)

    # Socio-political aspect
    if optim_scenario == 3 or optim_scenario == 5:
        model.NPT_constratint = Constraint(model.i, model.t, rule=countries_NPT)
        model.nuclear_perception_constraint = Constraint(model.i, model.t, rule=public_perception_constraint)
        model.ccs_included_constraint = Constraint(model.i, model.t, rule=ccs_included_constraint)

    # Technical aspect
    if optim_scenario == 4 or optim_scenario == 5:
        model.water_reserve_constraint = Constraint(model.i, model.t, rule=water_reserve_constraint)
        model.land_reserve_constraint = Constraint(model.i, model.t, rule=land_reserve_constraint)  

    print("Model construction finished.")
    #%% Solver

    print("Solving the model...")
    # Solver

    solver_name = "highs"
    solver = SolverFactory(solver_name + "nl",
                            executable=modules.find(solver_name),
                            solve_io="nl")
    result = solver.solve(model, tee=True)

    print(f'The solver status is {result.solver.status}')
    print(f'The termination condition is {result.solver.termination_condition}')

    if result.solver.termination_condition == 'infeasible':
        print("The model is infeasible. Please check the constraints and data.")
        log_infeasible_constraints(model)

    #%% Result export
    single_DACCS_capacity = 1e9/1e6  # 1 plant capacity is 1 million ton CO2/year
    single_nuclear_capacity = 1000/300  # 1 plant capacity is 300 MW

    if result.solver.termination_condition == 'optimal':
        print(f'the total present cost is: {model.obj():.2f} Trillion USD')

        sum_export = {
        (i, t): sum(value(model.export_elc[i, t, h]) for h in model.h)
        for i in model.i
        for t in model.t
        }

        sum_excess = {
        (i, t): sum(value(model.excess_electricity[i, t, h]) for h in model.h)
        for i in model.i
        for t in model.t
        }
        
        # Getting result
        result_dict = {
        "New Capacity": np.zeros((num_countries, num_periods)),
        "Capacity": np.zeros((num_countries, num_periods)),
        "Number of DACCS plant": np.zeros((num_countries, num_periods)),
        "New Nuclear Capacity": np.zeros((num_countries, num_periods)),
        "Nuclear Capacity": np.zeros((num_countries, num_periods)),
        "Number of Nuclear plant": np.zeros((num_countries, num_periods)),
        "Investment Required": np.zeros((num_countries, num_periods)),
        "Investment Required as percent GDP": np.zeros((num_countries, num_periods)),
        "Export electricity": np.zeros((num_countries, num_periods)),
        "Excess electricity": np.zeros((num_countries, num_periods)),
        "Land consumption": np.zeros((num_countries, num_periods)),
        "Water consumption": np.zeros((num_countries, num_periods)),
        "CO2 Storage Level": np.zeros(num_countries),
        "CO2 ppm": np.zeros(num_periods),
        "Objective Value": pd.DataFrame(
        {"Objective Value": [float(model.obj())]})
        }

        for t in range(num_periods):
            result_dict[f'cf NDACS {years[t]}'] = np.zeros((num_countries, num_hours))               # co2 concentration in atmosphere in ppm

        for x, i in enumerate(model.i):
            for y, t in enumerate(model.t):
                result_dict['New Capacity'][x,y] = model.new[i, t].value          # New capacity in million CO2/year
                result_dict['Capacity'][x, y] = model.capacity[i, t]()       # Capacity in Gt CO2/year
                result_dict['Number of DACCS plant'][x, y] = int(model.capacity[i, t]()*single_DACCS_capacity)      
                result_dict['New Nuclear Capacity'][x,y] = model.new_nuclear[i, t].value          # New capacity in million CO2/year
                result_dict['Nuclear Capacity'][x, y] = model.nuclear_capacity[i,t]()       # Capacity in million CO2/year
                result_dict['Number of Nuclear plant'][x, y] = int(model.nuclear_capacity[i,t]()*single_nuclear_capacity)       # Capacity in million CO2/year
                result_dict['Investment Required'][x,y] = value((model.capex[t] * 1e9 * model.new[i,t]) + (model.nuclear_capex[t] * 1e6 * model.new_nuclear[i,t])) / 1e9   # Investment required in billion USD
                result_dict['Investment Required as percent GDP'][x,y] = result_dict['Investment Required'][x,y] * 1e9 / (model.gdp[i,t] * 1e6) * 100   # Investment required as percent of GDP
                result_dict['Export electricity'][x,y] = sum_export[(i, t)] / 1000 /3600     # Excess electricity in TWh
                result_dict['Excess electricity'][x,y] = sum_excess[(i, t)] / 1000 /3600     # Excess electricity in TWh
                result_dict['Land consumption'][x,y] = model.land_consumption[i, t]() / model.land_reserve[i]      # Land consumption in km2
                result_dict['Water consumption'][x,y] = model.water_consumption[i, t]() / model.water_reserve[i]        # Water consumption in billion m3
                
                if y == 0:
                    result_dict['CO2 Storage Level'][x] = sum(model.captured_co2[i,t, h]() for t in model.t for h in model.h) / max(1e-10, model.co2_storage[i])       # co2 storage level (100% means full 0% empty)

                if x == 0:
                    result_dict['CO2 ppm'][y] = model.co2_ppm[t]()               # co2 concentration in atmosphere in ppm
                
                for z, h in enumerate(model.h):
                    try:
                        result_dict[f'cf NDACS {t}'][x,z] = model.captured_co2[i,t,h].value / interval[t] / model.capacity[i, t]()    # capacity factor of Nuclear DACCS
                    except(ZeroDivisionError):
                        result_dict[f'cf NDACS {t}'][x,z] = 0

        for key in ["New Capacity", "Capacity", "New Nuclear Capacity", 
                    "Nuclear Capacity", "Investment Required", "Export electricity", 
                    "Number of DACCS plant", "Number of Nuclear plant", "Excess electricity", 
                    "Land consumption", "Water consumption", "Investment Required as percent GDP"]:      # # data that belong to i and t set
            result_dict[key] = np.column_stack((countries, result_dict[key]))  
            result_dict[key] = pd.DataFrame(result_dict[key], columns=['Country'] + list(years))

        for key in ["CO2 Storage Level"]:     # data that only belong to i set
            result_dict[key] = np.column_stack((countries, result_dict[key]))  
            result_dict[key] = pd.DataFrame(result_dict[key], columns=['Country', key])

        for key in ["CO2 ppm"]:     # data that only belong to t set
            result_dict[key] = np.column_stack((years, result_dict[key]))  
            result_dict[key] = pd.DataFrame(result_dict[key], columns=['Year', key])

        for t in range(num_periods):        # data that belong to i and h set
            key = f'cf NDACS {years[t]}'
            result_dict[key] = np.column_stack((countries, result_dict[key]))  
            result_dict[key] = pd.DataFrame(result_dict[key], columns=['Country'] + [f'Hour {h}' for h in range(num_hours)])

        with pd.ExcelWriter(f"Result/Result {case_name} {co2_scenario}.xlsx") as writer:
            for sheet_name, df in result_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print("Excel file created successfully!")
    
    end = time.time()

    print(f'total runtime: {end-start:.2f} s')
    return model