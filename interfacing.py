import matplotlib.pyplot as pyplt
import py_dss_interface
import pandas as pd
data=pd.read_csv(r"C:\Users\anup.parajuli\Desktop\pythonProject\1440_similar.csv")
dss=py_dss_interface.DSSDLL(r"C:\Program Files\OpenDSS")
dss_file=r"C:\Users\anup.parajuli\Desktop\pythonProject\123_JPT.dss"
dss.text(f"compile [{dss_file}")
dss.solution_solve()
#v=dss.text("show voltages LN nodes")



