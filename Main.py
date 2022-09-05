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
            # Obtenemos los datos y contamos la frecuencia
            df = pd.read_excel(f, header=0)


if __name__ == "__main__":
    main()