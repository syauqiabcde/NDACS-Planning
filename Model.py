import numpy as np
from pyomo.environ import *
import pandas as pd
import openpyxl as xl 
import time

start = time.time()

# General data and asumption
ppm_to_mass = 2.13      # Conversion from ppm to mass CO2, 1 ppm of CO2 = 2.13 giga ton of CO2
plant_capacity_factor = 0.9
scenario = 'IPCC middle'
percent_allowable_investment = 0.05
plant_lifetime = 30
uranium_consumption = 1 # in kg uranium/capacity/year 
uranium_reserve = 8e9 # uranium reserve in the world
water_consumption = 1 # in m3/capacity/year
land_consumption = 1 # in m2/capacity

# Getting data from excel
wb = xl.load_workbook('data.xlsx') 
sheet_name = wb.sheetnames
df = {sheet: pd.read_excel('data.xlsx', sheet_name=sheet) for sheet in sheet_name}
years = df[sheet_name[0]].columns[1:]
countries = df[sheet_name[0]].iloc[:,0].values
nuclear_allowed = df['nuclear_allowed'].iloc[:,:].values.reshape(-1)
daccs_included = df['daccs_included'].iloc[:,:].values.reshape(-1)
interval = years[1] - years[0]

num_periods = len(years)
num_countries = len(countries)

# Define model
model = ConcreteModel()

# Sets
model.i = Set(initialize=countries)
model.i_nuclear_allowed = Set(initialize=nuclear_allowed)
model.i_daccs_included = Set(initialize=daccs_included)
model.t = Set(initialize=years)

# Parameters
cost_table = df['cost_table'].iloc[:,1:].values
co2_tax = df['co2_tax'].iloc[:,1:].values
co2_storage_capacity = df['co2_storage_capacity'].iloc[:,1].values
gdp = df['gdp'].iloc[:,1:].values
plant_capex = df['capex'].iloc[:,1].values
ccs_readiness_index = df['ccs_readiness'].iloc[:,1].values
nuclear_readiness_index = df['nuclear_readiness'].iloc[:,1].values
nuclear_perception = df['nuclear_perception'].iloc[:,1].values
water_reserve = df['resource_reserve'].iloc[:,1].values
land_reserve = df['resource_reserve'].iloc[:,2].values

cost_data = {(countries[i], years[t]): cost_table[i, t] for i in range(num_countries) for t in range(num_periods)}
co2_tax_data = {(countries[i], years[t]): co2_tax[i, t] for i in range(num_countries) for t in range(num_periods)}
CO2_storage_data = {countries[i]: co2_storage_capacity[i] for i in range(num_countries)}
gdp_data = {(countries[i], years[t]): gdp[i, t] for i in range(num_countries) for t in range(num_periods)}
capex_data = {years[t]: plant_capex[t] for t in range(num_periods)}
ccs_readiness_data = {countries[i]: ccs_readiness_index[i] for i in range(num_countries)}
nuclear_readiness_data = {countries[i]: nuclear_readiness_index[i] for i in range(num_countries)}
nuclear_perception_data = {countries[i]: nuclear_perception[i] for i in range(num_countries)}
water_reserve_data = {countries[i]: water_reserve[i] for i in range(num_countries)}
land_reserve_data = {countries[i]: land_reserve[i] for i in range(num_countries)}

model.cost_table = Param(model.i, model.t, initialize=cost_data, within=Reals)
model.co2_tax = Param(model.i, model.t, initialize=co2_tax_data, within=Reals)
model.CO2_storage = Param(model.i, initialize=CO2_storage_data , within=Reals)
model.gdp = Param(model.i, model.t, initialize=gdp_data , within=Reals)
model.capex = Param(model.t, initialize=capex_data , within=Reals)
model.ccs_readiness = Param(model.i, initialize=ccs_readiness_data , within=Reals)
model.nuclear_readiness = Param(model.i, initialize=nuclear_readiness_data , within=Integers)
model.nuclear_perception = Param(model.i, initialize=nuclear_perception_data , within=Reals)
model.water_reserve = Param(model.i, initialize=water_reserve_data , within=Reals)
model.land_reserve = Param(model.i, initialize=land_reserve_data , within=Reals)

# Decision Variables
model.new = Var(model.i, model.t, within=NonNegativeReals, bounds=(0.0, 100.0))
model.co2_ppm = Var(model.t, within=NonNegativeReals)       # not really a decision variables just a way to determine the co2 ppm based on the previous value of co2 ppm

# Expressions
def capacity_expansion(model, i, t):
    if t == min(model.t):
        return model.new[i,t]
    else:
        return model.capacity[i,t-interval] + model.new[i,t] - model.retired[i,t]

def retirement(model, i, t):
        retirement_period = t - plant_lifetime  # Find the period that corresponds to retirement period
        if retirement_period in model.t:
            return model.new[i, retirement_period]  # Retire what has been exceed lifetime
        else:
            return 0

def _IPCC_prediction(t, scenario='IPCC middle'):        # return the CO2 concentration in ppm based on IPCC prediction 
    if scenario == 'IPCC middle':
        return -0.00804 * t**2 + 34.83214 * t - 37159.42857
    elif scenario == 'IPCC conservative':
        return 0.05091 * t**2 - 203.18805 * t + 203146.33017 

def captured_co2(model, i, t):
    return model.capacity[i,t] * plant_capacity_factor

model.retired = Expression(model.i, model.t, rule=retirement)
model.capacity = Expression(model.i, model.t, rule=capacity_expansion)
model.captured_co2 = Expression(model.i, model.t, rule=captured_co2)

# Objective Function
def tot_cost(model):
    return sum((model.cost_table[i, t] - model.co2_tax[i, t]) * model.capacity[i, t] for i in model.i for t in model.t)

model.obj = Objective(rule=tot_cost , sense=minimize)

# Mandatory Constraints
def CO2_captured(model, t):
    if t == max(model.t):
        return model.co2_ppm[t] <= 411        # the co2 ppm in 2100 must lower than 411 ppm (Paris agreement 1.5 C scenario) 
    else:
        return model.co2_ppm[t] <= 465        # the co2 ppm in all years must lower than 465 ppm (Paris agreement 1.5 C scenario) 

def CO2_storage(model, i):
    return sum(model.captured_co2[i,t] for t in model.t) <= model.CO2_storage[i]

def co2_ppm_rule(model, t):
    if t == min(model.t):  # Initial CO2 concentration at the first period
        return model.co2_ppm[t] == _IPCC_prediction(t, scenario)
    else:
        generated_co2 = (_IPCC_prediction(t, scenario) - _IPCC_prediction(t-interval, scenario)) * ppm_to_mass
        captured = sum(model.captured_co2[i, t] for i in model.i)
        return model.co2_ppm[t] == model.co2_ppm[t-interval] + (generated_co2 - captured) / ppm_to_mass
    
model.capture_constraint = Constraint(model.t, rule=CO2_captured)
model.storage_constraint = Constraint(model.i, rule=CO2_storage)
model.co2_ppm_constraint = Constraint(model.t, rule=co2_ppm_rule)

# Optional Constraints (please comment in which constraint you dont want to apply)
def financing_constrant(model, i, t):
    return model.new[i, t] * model.capex[t] <= percent_allowable_investment * model.gdp[i, t]

def profit_constraint(model, i, t):
    return model.captured_co2[i,t] * model.co2_tax[i,t] >= model.capacity[i,t] * model.cost_table[i,t]

def ccs_readiness_constraint(model, i):
    if model.ccs_readiness[i] < 50:
        return model.new[i,t] == 0
    return Constraint.Skip

def nuclear_readiness_constraint(model, i):
    if model.nuclear_readiness[i] < 3:
        return model.new[i,t] == 0
    return Constraint.Skip

def countries_allowed_constraint(model, i, t):
    if i not in model.i_nuclear_allowed.data():
        return model.new[i,t] == 0
    return Constraint.Skip

def public_support_constraint(model, i):
    if model.nuclear_perception[i] < 50:
        return model.new[i,t] == 0
    return Constraint.Skip

def uranium_reserve_constraint(model):
    return sum(model.capacity[i,t] for i in model.i for t in model.t) * uranium_consumption <= uranium_reserve

def water_reserve_constraint(model, i):
    return sum(model.capacity[i,t] for t in model.t) * water_consumption <= model.water_reserve[i]

def land_reserve_constraint(model, i):
    return sum(model.capacity[i,t] for t in model.t) * land_consumption <= model.land_reserve[i]

def daccs_included_constraint(model, i):
    if i not in model.i_daccs_included:
        return model.new[i,t] == 0
    return Constraint.Skip

# Economic aspect
model.financing_constraint = Constraint(model.i, model.t, rule=financing_constrant)
# model.profitable_constraint = Constraint(model.i, model.t, rule=profit_constraint)

# Technological readiness
model.ccs_readiness_constraint = Constraint(model.i, rule=ccs_readiness_constraint)
model.nuclear_readiness_constraint = Constraint(model.i, rule=nuclear_readiness_constraint)

# Socio-political aspect
model.nuclear_allowed = Constraint(model.i, model.t, rule=countries_allowed_constraint)
model.nuclear_perception_constraint = Constraint(model.i, rule=public_support_constraint)

# Technical aspect
model.uranium_reserve_constraint = Constraint(rule=uranium_reserve_constraint)
model.water_reserve_constraint = Constraint(model.i, rule=water_reserve_constraint)
model.land_reserve_constraint = Constraint(model.i, rule=land_reserve_constraint)

# Political wilingness aspect
model.daccs_included_constraint = Constraint(model.i, rule=daccs_included_constraint)

# Solver
solver = SolverFactory('glpk', executable='C:\\w64\\glpsol.exe')
result = solver.solve(model)

print(f'The optimization status is {result.solver.status}')
print(f'The termination condition is {result.solver.termination_condition}')

# Getting result
capacity_array = np.zeros((num_countries, num_periods))
new_cap_array = np.zeros((num_countries, num_periods))
co2_ppm_array = np.zeros(num_periods)

for x, i in enumerate(model.i):
    for y, t in enumerate(model.t):
        new_cap_array[x,y] = model.new[i, t].value
        capacity_array[x, y] = model.capacity[i, t]()

        if x == 0:
            co2_ppm_array[y] = model.co2_ppm[t]()

new_cap_array = np.concatenate((countries.reshape(-1,1), new_cap_array), axis=1)
capacity_array = np.concatenate((countries.reshape(-1,1), capacity_array), axis=1)

end = time.time()

print(f'total runtime: {end-start:.2f} s')