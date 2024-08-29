import pandas as pd
from func_tools import cleaner_chain

df1 = pd.read_excel(
    "./NAMs_inventory_v20240723-A-L.xlsx",
    sheet_name="mynamms",
)
df2 = pd.read_excel(
    "./NAMs_inventory_v20240723_M-Z.xlsx",
    sheet_name="mynamms",
)
df2.columns = df1.columns
inventory = pd.concat([df1, df2], ignore_index=True)
cleaner_chain(inventory, "first_stage.xlsx")
print("Done!")
