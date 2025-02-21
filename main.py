from api.api import consultar_datos
from ui.ui import pedir_datos
from ui.ui import ordenar_tabular_datos

def main():
    cantidad_registros, nombre_departamento = pedir_datos()
    results_df = consultar_datos(cantidad_registros, nombre_departamento)
    ordenar_tabular_datos(results_df)
    
if __name__ == "__main__":
    main()

