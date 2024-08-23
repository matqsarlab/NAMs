import os

import pandas as pd
import openpyxl


def read_xlsx_paths(root_dir: str):
    return [
        os.path.join(dp, f)
        for dp, _, fnms in os.walk(root_dir)
        for f in fnms
        if f.endswith(".xlsx")
    ]


def concate():
    # WCZYTAJ PLIK
    main_table = pd.read_excel("TEMPLATE.xlsx", sheet_name="table")
    main_table = pd.read_excel("~/Downloads/TEMPLATE.xlsx", sheet_name="table")
    paths = read_xlsx_paths("../ECHA_TEMPLATE_TASK/Output/")
    paths = [
        "~/Downloads/Chen_2022_DNA-Oxidative-Damage-as-a-Sensitive-Genetic-Endpoint_CARC.xlsx",
        "~/Downloads/Dupa.xlsx",
        "~/Downloads/Dupa2.xlsx",
    ]

    number = 1
    for path in paths:
        try:
            exc = pd.read_excel(
                path, sheet_name=["SCORING", "NAMs", "S&K"], engine="openpyxl"
            )
            workbook = openpyxl.load_workbook(os.path.expanduser(path), data_only=True)
            nams_sheet = workbook["NAMs"]
            sk_sheet = workbook["S&K"]
            s_score = sk_sheet["F4"].value
            k_score = sk_sheet["F8"].value
            nams = pd.DataFrame(
                exc["NAMs"].iloc[4:, 0:].values, columns=exc["NAMs"].iloc[3].to_list()
            )
            nams["No"] = number
            nams["Authors"] = nams_sheet["A3"].value
            nams["Title"] = nams_sheet["B3"].value
            nams["Doi"] = nams_sheet["C3"].value
            nams["Evaluator"] = nams_sheet["D3"].value
            nams["Endpoint (1st level of description)"] = nams_sheet["E3"].value
            nams["Alternative endpoint 1 (2nd level of description)"] = nams_sheet[
                "F3"
            ].value
            nams["Alternative endpoint 2 (3rd level of description)"] = nams_sheet[
                "G3"
            ].value
            nams["S SCORE"] = s_score
            nams["K SCORE"] = k_score
            main_table = pd.concat([main_table, nams], ignore_index=True)
            # main_table[nams.columns] = nams.values
            # main_table["No"] = number
            number += 1

        except FileNotFoundError:
            print(f"File not found: {path}")
    main_table.dropna(axis=1, inplace=True)
    main_table.to_excel("to_remove.xlsx")


concate()
