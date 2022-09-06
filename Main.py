from multiprocessing.util import info
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

            # Listado de ministerios, localidades, puestos y género
            ministerios = df['MINISTERIO / ENTE / ORGANISMO'].unique()

            # Agrupamos por ministerio
            estadisticas_por_ministerio = df.groupby('MINISTERIO / ENTE / ORGANISMO')
            wb = openpyxl.Workbook()
            hoja = wb.active
            hoja['A1'].value = "Curso: " + archivo.split('.')[0]
            hoja['B1'].value = "Total: " + str(len(df.index))
            i = 5

            # TO-DO
            # Selección de campos
            # Estilos 
            # Merge de celdas

            # Sin género
            for ministerio in ministerios:
                estadisticas_ministerio = estadisticas_por_ministerio.get_group(ministerio).agg('value_counts')
                localidades = estadisticas_ministerio.reset_index(name='TOTAL')['LOCALIDAD']
                funciones =  estadisticas_ministerio.reset_index(name='TOTAL')['PUESTO ACTUAL']
                totales = estadisticas_ministerio.reset_index(name='TOTAL')['TOTAL']
                print(estadisticas_ministerio)
                print('--------------' + ministerio + '----------------')
                hoja['A' + str(i)].value = ministerio
                for localidad, funcion, total in zip(localidades, funciones, totales):
                    hoja['B' + str(i)].value = localidad
                    hoja['C' + str(i)].value = funcion
                    hoja['D' + str(i)].value = total
                    i += 1

            # Con género
            # for ministerio in ministerios:
            #      estadisticas_ministerio = estadisticas_por_ministerio.get_group(ministerio).agg('value_counts')
            #      localidades = estadisticas_ministerio.reset_index(name='TOTAL')['LOCALIDAD']
            #      funciones =  estadisticas_ministerio.reset_index(name='TOTAL')['PUESTO ACTUAL']
            #      generos = estadisticas_ministerio.reset_index(name='TOTAL')['GÉNERO']
            #      totales = estadisticas_ministerio.reset_index(name='TOTAL')['TOTAL']
            #      print('--------------' + ministerio + '----------------')
            #      hoja['A' + str(i)].value = ministerio
            #      for localidad, funcion, genero, total in zip(localidades, funciones, generos, totales):
            #         hoja['B' + str(i)].value = localidad
            #         hoja['C' + str(i)].value = funcion
            #         if genero == 'MASCULINO':
            #             hoja['D' + str(i)].value = total
            #             hoja['E' + str(i)].value = 0
            #         else:
            #             hoja['D' + str(i)].value = 0
            #             hoja['E' + str(i)].value = total
            #         i += 1
            
            wb.save(excel_output)
            excel_input = "Archivos\\Entrada"
            excel_output = "Archivos\\Salida"

if __name__ == "__main__":
    main()