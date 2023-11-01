import matplotlib.pyplot as pyplt
import py_dss_interface
import pandas as pd
import numpy as np
import os
import pathlib
import math
import openpyxl
data=pd.read_csv(r"C:\Users\anup.parajuli\Desktop\pythonProject\1440_similar.csv")
print(data)
# dss=py_dss_interface.DSSDLL(r"C:\Program Files\OpenDSS")
# dss_file=r"C:\Users\anup.parajuli\Desktop\pythonProject\prev_123JPT.dss"
script_path=os.path.dirname(os.path.abspath(__file__))
dss_file=pathlib.Path(script_path).joinpath("prev_123JPT.dss")
dss=py_dss_interface.DSSDLL()
interval=data.iloc[:,0]
print(interval)
loadshape=data.iloc[:,1]
dss.text(f"compile [{dss_file}]")
dss.text("New LoadShape.similar npts=1440 minterval=0 csvfile=1440_similar.csv")
# dss.text("New energymeter.meter element=Line.1_2 terminal=1")
dss.text("Batchedit Load..* daily=similar")
dss.text("set mode=daily")
dss.text("set number=1440")
dss.text("set stepsize=1m")
dss.solution_solve()
dss.text("show voltages LN")










