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
            localidades = df['LOCALIDAD'].unique()
            puestos = df['PUESTO ACTUAL'].unique()
            generos = df['GÉNERO'].unique()

            estadisticas = df.groupby('MINISTERIO / ENTE / ORGANISMO')

            for ministerio, grupo in estadisticas:
                por_localidad = estadisticas.get_group(ministerio).groupby('LOCALIDAD')
                print('-----------------------' + ministerio + '-----------------------')
                for localidad, grupo in por_localidad:
                    por_puesto_actual = por_localidad.get_group(localidad).groupby('PUESTO ACTUAL')
                    #print(por_localidad.get_group(localidad)['PUESTO ACTUAL'])
                    for puesto_actual, grupo in por_puesto_actual:
                        print(por_puesto_actual.get_group(puesto_actual)['PUESTO ACTUAL'])

            
            # Creamos las estadísticas
            # for ministerio in ministerios:
            #     informacion_ministerio = df.loc[df['MINISTERIO / ENTE / ORGANISMO'].isin([ministerio])]
            #     ministerio_localidad = informacion_ministerio['LOCALIDAD'].unique()
            #     ministerio_localidad_funcion = informacion_ministerio['LOCALIDAD'].loc[informacion_ministerio['PUESTO ACTUAL'].isin([puestos])].unique()

            #     #print(ministerio_localidad)
            #     estadisticas[ministerio]["ESTADISTICAS"] = {
            #         'LOCALIDADES': ministerio_localidad,
            #         'FUNCION': ministerio_localidad_funcion,

            #     }
            # print(estadisticas["ASIP"]["ESTADISTICAS"])
            #filtro_ministerio = df.loc[df['MINISTERIO / ENTE / ORGANISMO'].isin(['AMA'])]
            #localidades = filtro_ministerio['LOCALIDAD']
            #filtro_localidad = filtro_ministerio.loc[filtro_ministerio['PUESTO ACTUAL'].isin(['a'])]
            #print(filtro_localidad)
            #for t in test:
            #    print(t)
            #print(test)
            #for ministerio in ministerios:
            #    ministerios[ministerio]

if __name__ == "__main__":
    main()