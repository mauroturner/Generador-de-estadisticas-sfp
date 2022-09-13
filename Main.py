import numpy as np
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font
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
            hoja.column_dimensions['A'].width = 25
            hoja.column_dimensions['B'].width = 25
            hoja.column_dimensions['C'].width = 25

            # Generamos las estadísticas
            i = 3
            for ministerio, grupo in estadisticas:
                comienza_en_fila = i
                localidad_comienza_en_fila = i
                por_localidad = estadisticas.get_group(ministerio).groupby('LOCALIDAD')
                for localidad, grupo in por_localidad:
                    por_puesto_actual = por_localidad.get_group(localidad).groupby('PUESTO ACTUAL')
                    hoja['B' + str(i)].value = localidad
                    for puesto_actual, grupo in por_puesto_actual:
                        hoja['C' + str(i)].value = puesto_actual
                        hoja['C' + str(i)].alignment = Alignment(horizontal='center', vertical='center')
                        hoja['C' + str(i)].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                        hoja['C' + str(i)].font = Font(name='Arial', size=8)
                        # hoja['D' + str(i)].value = puesto_actual
                        hoja['D' + str(i)].alignment = Alignment(horizontal='center', vertical='center')
                        hoja['D' + str(i)].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                        hoja['D' + str(i)].font = Font(name='Arial', size=8)
                        i += 1
                    hoja.merge_cells('B' + str(localidad_comienza_en_fila) + ':' + 'B' + str(i-1))
                    hoja['B' + str(localidad_comienza_en_fila)].alignment = Alignment(horizontal='center', vertical='center')
                    hoja['B' + str(localidad_comienza_en_fila)].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    hoja['B' + str(localidad_comienza_en_fila)].font = Font(name='Arial', size=8)
                    hoja.merge_cells('E' + str(comienza_en_fila) + ':' + 'E' + str(i-1))
                    hoja['E' + str(comienza_en_fila)].alignment = Alignment(horizontal='center', vertical='center')
                    hoja['E' + str(comienza_en_fila)].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    hoja['E' + str(comienza_en_fila)].font = Font(name='Arial', size=8)
                    localidad_comienza_en_fila = i
                hoja.merge_cells('A' + str(comienza_en_fila) + ':' + 'A' + str(i-1))
                hoja['A' + str(comienza_en_fila)].alignment = Alignment(horizontal='center', vertical='center')
                hoja['A' + str(comienza_en_fila)].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                hoja['A' + str(comienza_en_fila)].font = Font(name='Arial', size=8)
                hoja['A' + str(comienza_en_fila)].value = ministerio
            wb.save(os.path.join('Archivos\\Salida', archivo))
        # limpiamos la ruta
        f = ''

if __name__ == '__main__':
    main()