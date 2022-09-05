import numpy as np
import openpyxl
import pandas as pd
import os 

"""
Estadísticas por 
- Ministerio
- Localidad
- Puesto actual
- Género
"""

def main():
    excel_input = "Archivos\\Entrada"
    excel_output = "Archivos\\Salida"
    for archivo in os.listdir(excel_input):
        f = os.path.join(excel_input, archivo)
        excel_output = os.path.join(excel_output, archivo)
        if os.path.isfile(f):
            # Obtenemos los datos y ordenamos
            df = pd.read_excel(f, header=0).sort_values('MINISTERIO / ENTE / ORGANISMO')
            #print(df.sort_values('MINISTERIO / ENTE / ORGANISMO'))
            # Creamos un diccionario de ministerios
            ministerios = dict.fromkeys(df['MINISTERIO / ENTE / ORGANISMO'].unique(), {'LOCALIDADES': 'rio gallegos'})
            #test = df.loc[df['MINISTERIO / ENTE / ORGANISMO'].isin(['ASIP'])]
            #print(test)
            for ministerio in ministerios:
                print(ministerio)

if __name__ == "__main__":
    main()