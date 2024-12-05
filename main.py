import os

import pandas as pd
import dbf
from dbf import Table

folder_path = 'archivos'
excel_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xlsx')]

for excel_file in excel_files:
    df = pd.read_excel(excel_file, engine='openpyxl', dtype={'Dni': str})
    df = df.replace({'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U', 'Ü': 'U', '-': ' ', ',': ' '},
                    regex=True)
    df = df.rename(columns={
        'firstName': 'NOM_ADE',
        'lastNameP': 'APE_PAT',
        'lastNameM': 'APE_MAT',
        'Dni': 'NUM_ELE',
        'serieJNE': 'NUM_PAG'
    })
    df['NOM_ADE'] = df['NOM_ADE'].astype(str)
    df['APE_PAT'] = df['APE_PAT'].astype(str)
    df['APE_MAT'] = df['APE_MAT'].astype(str)
    df['NUM_ELE'] = df['NUM_ELE'].astype(str)
    df['NUM_PAG'] = pd.to_numeric(df['NUM_PAG'], errors='coerce')
    df['NUM_ITE'] = 1
    df = df.reindex(columns=['NUM_PAG', 'NUM_ITE', 'NUM_ELE', 'APE_PAT', 'APE_MAT', 'NOM_ADE'])
    file_name = excel_file.split('/')[-1].split('.')[0]
    table = Table(
        filename=file_name + '.dbf',
        codepage='cp1252',
        field_specs='NUM_PAG N(6,0); NUM_ITE N(2,0); NUM_ELE C(8); APE_PAT C(40); APE_MAT C(40); NOM_ADE C(35)'
    )
    table.open(dbf.READ_WRITE)
    for row in df.itertuples(index=False):
        table.append(row)
    table.close()
    print(f"Archivo DBF '{file_name}.dbf' creado correctamente.")