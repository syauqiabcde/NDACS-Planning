import numpy as np
from pyomo.environ import *

# Define model
model = ConcreteModel()

# Sets
years = np.arange(2030, 2101, 5)  # Years from 2030 to 2100
countries = ['Country1', 'Country2', 'Country3']  # Replace with actual country names

model.i = Set(initialize=countries)
model.t = Set(initialize=years)

#Parameters
num_periods = len(years)
num_countries = len(countries)
cost_table = np.random.uniform(1, 10, (len(countries), len(years)))
co2_tax = np.random.uniform(1, 10, (len(countries), len(years)))

# Parameters (convert numpy arrays into Pyomo Params)
cost_data = {(countries[i], years[t]): cost_table[i, t] for i in range(num_countries) for t in range(num_periods)}
co2_tax_data = {(countries[i], years[t]): co2_tax[i, t] for i in range(num_countries) for t in range(num_periods)}

model.cost_table = Param(model.i, model.t, initialize=cost_data, within=Reals)
model.co2_tax = Param(model.i, model.t, initialize=co2_tax_data, within=Reals)

# Decision Variables
model.capacity = Var(model.i, model.t, within=NonNegativeReals)

def tot_cost(model):
    return sum((model.cost_table[i, t] - model.co2_tax[i, t]) * model.capacity[i, t] for i in model.i for t in model.t)

model.obj = Objective(rule=tot_cost , sense=minimize)
solver = SolverFactory('glpk')
result = solver.solve(model)
