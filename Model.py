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
CO2_storage = np.random.uniform(1, 30, len(countries))

# Parameters (convert numpy arrays into Pyomo Params)
cost_data = {(countries[i], years[t]): cost_table[i, t] for i in range(num_countries) for t in range(num_periods)}
co2_tax_data = {(countries[i], years[t]): co2_tax[i, t] for i in range(num_countries) for t in range(num_periods)}
CO2_storage_data = {countries[i]: CO2_storage[i] for i in range(num_countries)}

model.cost_table = Param(model.i, model.t, initialize=cost_data, within=Reals)
model.co2_tax = Param(model.i, model.t, initialize=co2_tax_data, within=Reals)
model.CO2_storage = Param(model.i, initialize=CO2_storage_data , within=Reals)

def tot_cost(model):
    return sum((model.cost_table[i, t] - model.co2_tax[i, t]) * model.capacity[i, t] for i in model.i for t in model.t)

def CO2_captured(model, t):
    return sum(model.capacity[i,t] * 0.9 for i in model.i) >= 3.0

def CO2_storage(model, i):
    return sum(model.capacity[i,t] * 0.9 for t in model.t) <= model.CO2_storage[i]

def capacity_expansion(model, i, t):
    if t == min(model.t):
        return model.new[i,t]
    else:
        return model.capacity[i,t-5] + model.new[i,t] - model.retired[i,t]

def retirement(model, i, t):
        retirement_period = t - 30  # Find the period that corresponds to 30 years ago
        if retirement_period in model.t:
            return model.new[i, retirement_period]  # Retire what was built 30 years ago
        else:
            return 0

# Decision Variables
model.new = Var(model.i, model.t, within=NonNegativeReals, bounds=(0.0, 10.0))

# Expressions
model.retired = Expression(model.i, model.t, rule=retirement)
model.capacity = Expression(model.i, model.t, rule=capacity_expansion)

# Objective Function
model.obj = Objective(rule=tot_cost , sense=minimize)

# Mandatory Constraints
model.capture_constraint = Constraint(model.t, rule=CO2_captured)
model.storage_constraint = Constraint(model.i, rule=CO2_storage)

# Optional Constraints

solver = SolverFactory('glpk', executable='C:\\w64\\glpsol.exe')
result = solver.solve(model)

print(f'The optimization status is {result.solver.status}')
print(f'The termination condition is {result.solver.termination_condition}')

capacity_array = np.zeros((num_countries, num_periods))
new_cap_array = np.zeros((num_countries, num_periods))

for x, i in enumerate(model.i):
    for y, t in enumerate(model.t):
        new_cap_array[x,y] = model.new[i, t].value
        capacity_array[x, y] = model.capacity[i, t]()