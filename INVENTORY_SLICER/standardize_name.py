import pandas as pd

df1 = pd.read_excel("./template.xlsx")
inventory = pd.read_excel(
    "~/Downloads/NAMs_inventory_v20240723-A-L.xlsx",
    sheet_name="mynamms",
)
df2 = pd.read_excel(
    "~/Downloads/NAMs_inventory_v20240723_M-Z.xlsx",
    sheet_name="mynamms",
)
df2.columns = df1.columns
inventory = pd.concat([df1, df2], ignore_index=True)


def remove_blank(df):
    result = []
    columns = df.columns
    for _, row in df.iterrows():
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
    return pd.DataFrame(result, columns=columns)


def fill_missing_values(df):
    last_valid_row = None
    for index, row in df.iterrows():
        if pd.notna(row["No"]) and pd.notna(row["Authors"]):
            last_valid_row = row
        elif pd.isna(row["Authors"]):
            if last_valid_row is not None:
                for col in df.columns:
                    if pd.isna(row[col]):
                        df.at[index, col] = last_valid_row[col]
    return df


def fill_missing_values2(df):
    last_values = {}

    for index, row in df.iterrows():
        author = row["Authors"]

        if author not in last_values:
            last_values[author] = {}

        for col in df.columns:
            if col != "Authors" and pd.isna(row[col]):
                if col in last_values[author]:
                    df.at[index, col] = last_values[author][col]
            else:
                last_values[author][col] = row[col]

    return df


def remove_zeros(df, start=0, end=8):
    last_values = {}

    for index, row in df.iterrows():
        author = row["Authors"]

        if author not in last_values:
            last_values[author] = {}

        for col in df.columns[start:end]:
            if col != "Authors" and row[col] == 0:
                if col in last_values[author]:
                    df.at[index, col] = last_values[author][col]
            else:
                last_values[author][col] = row[col]

    df.iloc[:, start:end] = df.iloc[:, start:end].replace(0, pd.NA)
    return df


def cleaner_chain(inventory: pd.DataFrame, filename: str):
    result_df = remove_blank(inventory)
    result_df = result_df.dropna(subset=list(result_df.columns[10:]), how="all")
    result_df = fill_missing_values(result_df)
    result_df = fill_missing_values2(result_df)
    result_df = remove_zeros(result_df)
    return result_df.to_excel(filename, index=False)


cleaner_chain(inventory, "dupa4.xlsx")
