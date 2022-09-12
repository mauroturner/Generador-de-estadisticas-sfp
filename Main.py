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
    excel_input = 'Archivos\\Entrada'
    excel_output = 'Archivos\\Salida'
    for archivo in os.listdir(excel_input):
        f = os.path.join(excel_input, archivo)
        excel_output = os.path.join(excel_output, archivo)
        if os.path.isfile(f):
            # Obtenemos los datos y ordenamos
            df = pd.read_excel(f, header=0).sort_values('MINISTERIO / ENTE / ORGANISMO')

            # Agrupamos por ministerio
            estadisticas = df.groupby('MINISTERIO / ENTE / ORGANISMO')

            # Preparamos la hoja de cálculo
            wb = openpyxl.Workbook()
            hoja = wb.active
            hoja['A1'].value = "Curso: " + archivo.split('.')[0]
            hoja['B1'].value = 'Total: ' + str(len(df.index))

            # Generamos las estadísticas
            i = 3
            for ministerio, grupo in estadisticas:
                comienza_en_fila = i
                por_localidad = estadisticas.get_group(ministerio).groupby('LOCALIDAD')
                for localidad, grupo in por_localidad:
                    por_puesto_actual = por_localidad.get_group(localidad).groupby('PUESTO ACTUAL')
                    hoja['B' + str(i)].value = localidad
                    for puesto_actual, grupo in por_puesto_actual:
                        hoja['C' + str(i)].value = puesto_actual
                        i += 1
                hoja.merge_cells('A' + str(comienza_en_fila) + ':' + 'A' + str(i-1))
                hoja['A' + str(comienza_en_fila)].value = ministerio
            wb.save(os.path.join('Archivos\\Salida', archivo))
        # limpiamos la ruta
        f = ''

if __name__ == '__main__':
    main()