import os
import pandas as pd


file = 'Modelo\\959421.xlsx'

print(os.path.isfile(file))


df = pd.read_excel(file, sheet_name='Sheet1')
            
lista_etiquetas = df.values.tolist()

print("::::: Lista de etiquetas extraidas :::::")
for item in lista_etiquetas:
    print(item)