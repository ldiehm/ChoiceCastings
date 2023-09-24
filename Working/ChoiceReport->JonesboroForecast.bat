python3
from dataTransfer import Application
dateIndex = [1,2]
app = Application(appName= "Jonesboro Forecast Data Transfer", sourceFile= "Choice Report", sourceFileColumns= ["ICPROD","Total On Hand","On Order","Consigned","Required","SumOfWIP"], targetFile= "Jonesboro Forecast", targetFileColumns=["Part Number","CURRENT Inventory","CURR. Orders","KBM Cons.","CURRENT REQMNT.", "WIP"], dateLocation = dateIndex)
app.run()
quit()
