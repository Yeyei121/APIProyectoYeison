import pandas as pd
from sodapy import Socrata

def consultar_datos(limite_registros, nombre_departamento):
    client = Socrata("www.datos.gov.co", None)
    columnas_interes = 'ciudad_municipio_nom, departamento_nom, edad, fuente_tipo_contagio, estado, pais_viajo_1_nom'
    results = client.get(
    "gt2j-8ykr",
    limit=limite_registros,
    departamento_nom=nombre_departamento,
    select=columnas_interes)
    results_df = pd.DataFrame.from_records(results)
    return results_df

