import pandas as pd
from func_tools import pipleine2excels

df1 = pd.read_excel(
    "./NAMs_inventory_v20240723-A-L.xlsx",
    sheet_name="mynamms",
)
df2 = pd.read_excel(
    "./NAMs_inventory_v20240723_M-Z.xlsx",
    sheet_name="mynamms",
)
df1 = pd.read_excel(
    "~/Downloads/OneDrive_1_29-08-2024/AL.xlsx",
    sheet_name="Sheet1",
)
df2 = pd.read_excel(
    "~/Downloads/OneDrive_1_29-08-2024/MZ.xlsx",
    sheet_name="Sheet1",
)
# Wytnij pierwsze 33 kolumny:
df1 = df1.iloc[:, :33]
df2 = df2.iloc[:, :33]

df2.columns = df1.columns
inventory = pd.concat([df1, df2], ignore_index=True)
pipleine2excels(inventory)
print("Done!")
