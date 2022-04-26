import pandas as pd

# Converting the ISU_Tuition_10Yrs.xlsx into csv

# Read and store content of an excel file
read_file = pd.read_excel("ISU_Tuition_10Yrs.xlsx")

# Write the dataframe object into csv file
read_file.to_csv("ISU_Tuition_10Yrs.csv",
                 index=None,
                 header=True)

print(read_file)

# Credits for some of the source code to GeeksforGreeks.org
# URL: https://www.geeksforgeeks.org/convert-excel-to-csv-in-python/
