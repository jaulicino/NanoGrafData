import numpy as np
import pandas as pd

print("Enter the file path of the excel workbook you would like to pull from. If that workbook is in this directory, something like 'data_sample.xlsx' will suffice")
file_name = input()
print("Reading data from " + file_name+ " . . .      note: this may take between 30 seconds and a minute based on the size of the workbook")
df = pd.ExcelFile(file_name)
sheetnames = df.sheet_names
statistics = {}

'''
SECTION 1: Reads the statistics info from "Statistics*" tab of the workbook

'''
for i in sheetnames:
    if "Statistics" in i:
        statistics[i] = pd.read_excel(df, i)
requested_columns = ["Cycle_Index", "Charge_Capacity(Ah)", "Discharge_Capacity(Ah)"]
for i in statistics:
    sdf = pd.DataFrame(statistics[i][requested_columns], columns=requested_columns)
    sdf.name = i    
    sdf.to_csv(i+".csv", index=False)

'''
SECTION 2: Reads the channel info from "Channel*" tab of the workbook

'''
channels = {}
for i in sheetnames:
    if "Channel" in i:
        channels[i] = pd.read_excel(df, i)
requested_columns = ["Data_Point","Cycle_Index", "Current(A)", "Voltage(V)","Charge_Capacity(Ah)", "Discharge_Capacity(Ah)"]
for i in channels:
    cdf = pd.DataFrame(channels[i][requested_columns], columns=requested_columns)
    cdf.name = i    
    cdf.to_csv(i+".csv", index=False)

print("Read and Written Successfully")