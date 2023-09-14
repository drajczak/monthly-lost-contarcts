import pandas as pd
import os

# Load xlsx file
xlsx_path = r'K:\Mój dysk\Arkusze\Analiza Rynku Maszyn Budowlanych\_Analiza rynku maszyn budowlanych 2023-08.xlsx'
df = pd.read_excel(xlsx_path, sheet_name='dane', usecols=('kategoria', 'sekcja', 'wojewodztwo',	'okres', 'liczba'))

# Polish name of the month
df['okres'] = 'sierpień 2023'

# Delete rows with "kategoria" column containing specific words
categories_to_remove = ['STABILIZATORY', 'FREZARKI', 'ROZŚCIEŁACZE', 'WALCE', 'RÓWNIARKI', 'SPYCHARKI', 'TELESKOPOWE', 'KOPARKO ŁADOWARKI', 'MINI', 'SZTYWNORAMOWE']
df = df[~df['kategoria'].str.contains('|'.join(categories_to_remove))]

# Removing lines from the category 'ŁADOWARKI KOŁOWE' from section 0 - 120 HP
df = df[~((df['kategoria'] == 'ŁADOWARKI KOŁOWE  (<150KM)') & ((df['sekcja'] == '0 < 30 KM') | (df['sekcja'] == '30 < 60 KM') | (df['sekcja'] == '60 < 70 KM') | (df['sekcja'] == '70 < 80 KM') | (df['sekcja'] == '80 < 100 KM') | (df['sekcja'] == '100 < 120 KM'))) ]

#Deleting rows with column "liczba" equal to 0
df = df[df['liczba'] != 0]

# Duplicate rows for values greater than 1 in the "liczba" column
while True:
    mask = df['liczba'] > 1
    if not mask.any():
        break
    new_rows = df[mask].copy()
    new_rows['liczba'] = 1
    df.loc[mask, 'liczba'] -= 1
    df = pd.concat([df, new_rows], ignore_index=True)

# Creating new columns
df['PH'] = ''
df['Maszyna'] = ''
df['Klient'] = ''
df['Uwagi'] = ''

# Creating a new dictionary with the initials of traders and their regions
sales_regions = {
    'JKU': ['LUBUSKIE', 'ZACHODNIOPOMORSKIE', 'WIELKOPOLSKIE'],
    'PKO': ['LUBELSKIE', 'PODKARPACKIE', 'ŚWIĘTOKRZYSKIE'],
    'AKU': ['ŚLĄSKIE', 'OPOLSKIE'],
    'GPA': ['ŁÓDZKIE', 'ŚLĄSKIE'],
    'JKR': ['ŁÓDZKIE', 'MAZOWIECKIE'],
    'MKI': ['WARMIŃSKO_MAZURSKIE', 'KUJAWSKO_POMORSKIE', 'POMORSKIE'],
    # 'MPI': ['DOLNOŚLĄSKIE', 'LUBUSKIE'],
    'PKL': ['MAŁOPOLSKIE'],
    'PSZ': ['PODLASKIE', 'WARMIŃSKO_MAZURSKIE'],
    'OCZ': ['ZACHODNIOPOMORSKIE', 'WIELKOPOLSKIE','LUBELSKIE', 'PODKARPACKIE', 'ŚWIĘTOKRZYSKIE', 'ŚLĄSKIE', 'OPOLSKIE', 'ŁÓDZKIE', 'MAZOWIECKIE', 'DOLNOŚLĄSKIE', 'LUBUSKIE', 'MAŁOPOLSKIE', 'PODLASKIE', 'WARMIŃSKO_MAZURSKIE'],
    'JZA': ['ZACHODNIOPOMORSKIE', 'WIELKOPOLSKIE','LUBELSKIE', 'PODKARPACKIE', 'ŚWIĘTOKRZYSKIE', 'ŚLĄSKIE', 'OPOLSKIE', 'ŁÓDZKIE', 'MAZOWIECKIE', 'DOLNOŚLĄSKIE', 'LUBUSKIE', 'MAŁOPOLSKIE', 'PODLASKIE', 'WARMIŃSKO_MAZURSKIE']
}

# Check if the destination folder exists/create a folder
folder_path = r"K:/Mój dysk/Arkusze/Analiza Rynku Maszyn Budowlanych/DOSTARCZONE/2023-08/PUSTE"

if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# We create a dictionary with new DataFrames
new_dfs = {}
for key, regions in sales_regions.items():
    # Select rows according to the values in the 'województwo' column from the dictionary
    new_df = df[df['wojewodztwo'].isin(regions)].copy()
    # We complete the 'PH' column with the value of the key from the dictionary;
    new_df.loc[:, 'PH'] = key
    # We add a new DataFrame to the dictionary using the key from the dictionary
    new_dfs[key] = new_df

    # We create an Excel file for a given DataFrame
    filename = f'{folder_path}/2023-08_DOSTARCZONE_{key}.xlsx'
    with pd.ExcelWriter(filename) as writer:
        new_df.to_excel(writer, index=False)


# Saving a file
output_path = r'K:/Mój dysk/Arkusze/Analiza Rynku Maszyn Budowlanych/DOSTARCZONE/2023-08/2023-08 DOSTARCZONE.xlsx'
df.to_excel(output_path, index=False)
