import pandas as pd

inventory = pd.read_excel("./template.xlsx")


def fill_missing_values(df):
    last_valid_row = None
    col_num = 10
    for index, row in df.iterrows():
        if pd.notna(row["No"]) and pd.notna(row["Authors"]):
            last_valid_row = row
        elif pd.isna(row["Authors"]):
            if last_valid_row is not None:
                # df.loc[index] = last_valid_row
                for col in df.columns:
                    if col_num < 10:
                        df.at[index, col] = last_valid_row[col]
                        col_num += 1
                    elif pd.isna(row[col]):
                        df.at[index, col] = last_valid_row[col]

    return df


result = []
for index, row in inventory.iterrows():
    all_columns = row.values.copy()
    col20_value = all_columns[20]

    if isinstance(col20_value, str):
        elements = col20_value.split("\n")
        for element in elements:
            new_row = list(all_columns)
            new_row[20] = element
            result.append(new_row)
    else:
        result.append(all_columns)

result_df = pd.DataFrame(result, columns=inventory.columns)

result_df = result_df.dropna(subset=list(result_df.columns[10:]), how="all")

result_df = fill_missing_values(result_df)

result_df.to_excel("dupa.xlsx", index=False)
