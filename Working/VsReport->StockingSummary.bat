python3
from dataTransfer import Application
dateIndex = (1,4)
app = Application(appName= "VS Stocking Summary Data Transfer", sourceFile= "VS Report", sourceFileColumns= ["ICPROD", "Total On Hand","Consigned", "Required", "SumOfWIP"], targetFile= "Stocking Summary", targetFileColumns=["T&B     Part #", "Total Inv:  JBS & T&B", "VS Cons  On Hand", "Current Demand", "WIP"], dateLocation = dateIndex)
app.run()
quit()
