import matplotlib.pyplot as pyplt
import py_dss_interface
import pandas as pd
import numpy as np
import os
import pathlib
import math
import openpyxl
data=pd.read_csv(r"C:\Users\anup.parajuli\Desktop\pythonProject\1440_dissimilar_csv.csv")
print(data.shape)
interval=data.iloc[:,0]

# dss=py_dss_interface.DSSDLL(r"C:\Program Files\OpenDSS")
# dss_file=r"C:\Users\anup.parajuli\Desktop\pythonProject\prev_123JPT.dss"
script_path=os.path.dirname(os.path.abspath(__file__))
dss_file=pathlib.Path(script_path).joinpath("prev_123JPT.dss")
dss=py_dss_interface.DSSDLL()
#loadshape = np.zeros((data.shape[0], data.shape[1]))
loadshapes=[]
for i in range(data.shape[1]):
   loadshapes.append(data.iloc[:,i])
print(loadshapes[0])
dss.text(f"compile [{dss_file}]")
loads = dss.circuit_all_bus_names()
#dss.text("New LoadShape.dissimilar npts=1440 minterval=1 csvfile=1440_dissimilar_csv.csv")
#
loads = dss.circuit_all_bus_names()
print(len(loads))
for i in range(len(loads)):
    print(loads[i])
    dss.text(f"New LoadShape.loadshape{loads[i]} npts=1440 minterval=1 mult = (file=1440_dissimilar_csv.csv, col={loads[i]}, header=yes)")
    dss.text(f"New Monitor.monitor{loads[i]} element=load.{loads[i]} terminal=1")
    dss.text(f"edit load.{loads[i]} daily=loadshape{loads[i]}")

#dss.text("New LoadShape.similar npts=1440 minterval=0 csvfile=1440_similar.csv")

# dss.text("New energymeter.meter element=Line.1_2 terminal=1")
#dss.text("Batchedit Load..* daily=similar")

# for i in range(len(loads)):
#     dss.text(f"edit load.{i} daily=loadshape{[i]}")
# voltage=[]
dss.text("set mode=daily")
dss.text("set number=1440")
dss.text("set stepsize=1m")
dss.solution_solve()

# print(voltage)
dss.text("show voltages LN")
# voltage_data = pd.DataFrame()
#
# # Loop through each interval
# for interval in range(1440):
#     dss.text(f"set mode=daily")
#     dss.text(f"set number={interval + 1}")
#     dss.text("solve")
#
#     # Collect voltage data for each load
#     load_voltages = []
#     for i in range(len(loads)):
#         bus_name = loads[i]
#
#         # Set the active load element
#         dss.text(f"set active Load.{i+1}")
#
#         # Retrieve voltage magnitude for the active load
#         voltage = dss.Circuit.AllBusVmagPU()
#
#         load_voltages.append(voltage)
#
#     # Add the voltage data for the current interval to the DataFrame
#     voltage_data[f'Interval_{interval + 1}'] = load_voltages
#
# # Save the voltage data to an Excel file
# voltage_data.to_excel("voltages_data.xlsx", index=False)


















