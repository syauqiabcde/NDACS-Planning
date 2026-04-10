from robust_model_integrated import run_model
# from Model import RobustDACCSModel
from visualization import Plotting
import numpy as np  

co2_scenario=['IPCC middle']
ppm_limit_scenario = '1.5 C'
cases = [0,1,2,3,4,5]  
years = np.arange(2030, 2101, 5)

for scenario in co2_scenario:
    for case in cases:
        model = run_model(co2_scenario=scenario, 
                          ppm_limit_scenario=ppm_limit_scenario, 
                          optim_scenario=case)
        print('')
        plotter = Plotting(optim_scenario=case, co2_scenario=scenario)
        for y in years:
            plotter.plot_map(parameter='Capacity', year=y)
            plotter.plot_map(parameter='Land consumption', year=y, inpercent=True)
            plotter.plot_map(parameter='Water consumption', year=y, inpercent=True)

    plotter.plot_line(parameter='CO2 ppm', plot_cumsum= False)
    plotter.plot_line(parameter='Investment Required', plot_cumsum=True)
    plotter.plot_line(parameter='Capacity', plot_cumsum=False)
    plotter.plot_line(parameter='Nuclear Capacity', plot_cumsum=False)

plotter.plot_obj()