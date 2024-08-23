from scidownl import scihub_download
import pandas as pd
import re

paper_type = "doi"

# wczytanie pliku
endpoint = "./CARC_final.xlsx"
df = pd.read_excel(endpoint)


def extract_first_author(authors):
    # ekstrakcja pierwszego autora do zapisania nazwy pliku
    # Podzielić według pierwszego wystąpienia 'and' lub ','
    first_author = re.split(r" and |, ", authors)[0]
    # Usunąć wszystko po pierwszej spacji, łącznie ze spacją
    first_author = first_author.split(" ")[0]
    return first_author


def popraw_tytul(tytul):
    # ogarniecie tytułu do zapisu pliku
    tytul = re.sub(r'[<>:"/\\|?*]', "", tytul)  # Usunięcie niedozwolonych znaków
    tytul = re.sub(
        r"[\[\]{}().\-]", "", tytul
    )  # Usunięcie nawiasów, kropek i myślników
    tytul = re.sub(r"\s+", " ", tytul)  # Zamiana wielu spacji na pojedynczą spację
    return tytul.strip()


# Dodanie kolumny 'first_author' do DataFrame
df["first_author"] = df["authors"].apply(extract_first_author)
nazwa_endpointu = endpoint.replace(".xlsx", "")
print(nazwa_endpointu)
