def pedir_datos():
    cantidad_registros = int(input("Cantidad registros (máximo 1000): "))
    if cantidad_registros > 1000:
        print("Excedió el número de registros.")
        pedir_datos()
    nombre_departamento = input("Departamento (si ingresa una palabra no valida el programa no funcionará): ")
    nombre_departamento = nombre_departamento.upper()      
    return cantidad_registros, nombre_departamento


def filtrar_datos(results_df):
    if "pais_viajo_1_nom" not in results_df:
        results_df["pais_viajo_1_nom"] = "N/A"

    return results_df

def ordenar_tabular_datos(results_df):
    datos_filtrados = filtrar_datos(results_df)
    datos_filtrados = datos_filtrados.reset_index(drop=True)
    datos_filtrados.index += 1  
    datos_filtrados.insert(0, "Índice", datos_filtrados.index)
    nuevo_orden = ["Índice", "ciudad_municipio_nom", "departamento_nom", "edad", "fuente_tipo_contagio", "estado","pais_viajo_1_nom"]
    datos_filtrados = datos_filtrados[nuevo_orden]
    mapeo_columnas = {
        "Índice": "N°",
        "ciudad_municipio_nom": "Ciudad de ubicación", 
        "departamento_nom": "Departamento", 
        "edad": "Edad", 
        "fuente_tipo_contagio": "Tipo", 
        "estado": "Estado",
        "pais_viajo_1_nom": "País de procedencia"
    }
    datos_filtrados = datos_filtrados.rename(columns=mapeo_columnas)
    print("=" * 115)
    formato = "| {:^111} |"
    print(formato.format("Casos positivos de COVID-19 en Colombia"))
    print("=" * 115)
    nuevo_formato = "| {:<5} | {:<22} | {:<18} | {:<5} | {:<13} | {:<10} | {:<20} |"
    print(nuevo_formato.format("N°", "Ciudad de ubicación", "Departamento", "Edad", "Tipo", "Estado", "País de procedencia"))
    print("=" * 115)
    for index, fila in datos_filtrados.iterrows():
        print(nuevo_formato.format(
            fila["N°"], fila["Ciudad de ubicación"], fila["Departamento"], 
            fila["Edad"], fila["Tipo"], fila["Estado"], fila["País de procedencia"]
        ))
        print("-" * 115)
