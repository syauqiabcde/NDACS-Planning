import numpy as np
from pyomo.environ import *

# Define model
model = ConcreteModel()

# Sets
years = np.arange(2030, 2101, 5)  # Years from 2030 to 2100
countries = ['Country1', 'Country2']  # Replace with actual country names

model.i = Set(initialize=countries)
model.t = Set(initialize=years)

# Parameters (replace with actual data or lookup tables)
CAPEX = {t: 100 for t in years}  # Example CAPEX values
CRF = 0.07  # Capital Recovery Factor (example)
Fixed_OM = {t: 20 for t in years}  # Example fixed O&M cost
Uranium = {t: 5 for t in years}  # Example uranium cost
Uranium_disposal = {t: 3 for t in years}  # Example uranium disposal cost
transport = 10  # Transport cost (example)
storage = 5  # Storage cost (example)
carbon_tax = {(i, t): 50 for i in countries for t in years}  # Example carbon tax
CF = 0.9  # Capacity factor
CO2_storage_limit = {i: 1e6 for i in countries}  # Example storage limits
PA_1_5_requirement = 1e5  # Example Paris Agreement requirement

# Decision variables
model.new = Var(model.i, model.t, within=NonNegativeReals)

# Expressions
def annual_capex_rule(model, i, t):
    return CAPEX[t] * CRF * model.new[i, t]

def opex_rule(model, i, t):
    return Fixed_OM[t] + Uranium[t] + Uranium_disposal[t]

def unit_cost_rule(model, i, t):
    return (model.Annual_CAPEX[i, t] + model.OPEX[i, t]) / (model.capacity[i, t] + 1e-6) + transport + storage

# Objective function
def objective_rule(model):
    return sum(
        (model.unit_cost[i, t] - carbon_tax[i, t]) * model.capacity[i, t]
        for i in model.i for t in model.t
    )

# Constraints
def capacity_expansion_rule(model, i, t):
    if t == 2030:
        return model.capacity[i, t] == model.new[i, t]  # Initial condition
    return model.capacity[i, t] == model.capacity[i, t - 1] + model.new[i, t] - model.retired[i, t]

def co2_captured_rule(model, t):
    return sum(model.capacity[i, t] * CF for i in model.i) - PA_1_5_requirement >= 0

def co2_storage_rule(model, i):
    return sum(model.capacity[i, t] * CF for t in model.t) <= CO2_storage_limit[i]

model.cost = Objective(rule=objective_rule, sense=minimize)
model.Annual_CAPEX = Expression(model.i, model.t, rule=annual_capex_rule)
model.OPEX = Expression(model.i, model.t, rule=opex_rule)
model.unit_cost = Expression(model.i, model.t, rule=unit_cost_rule)

# Constraints
model.capacity_expansion = Constraint(model.i, model.t, rule=capacity_expansion_rule)
model.co2_captured = Constraint(model.t, rule=co2_captured_rule)
model.co2_storage = Constraint(model.i, rule=co2_storage_rule)

# Solve
solver = SolverFactory('glpk')
result = solver.solve(model, tee=True)

# Display results
for i in countries:
    for t in years:
        print(f"New capacity for {i} in {t}: {model.new[i, t].value}")
        print(f"Total capacity for {i} in {t}: {model.capacity[i, t].value}")
