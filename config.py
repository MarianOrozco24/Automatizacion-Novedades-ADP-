import json
import os
import sys

class Config:
    lista_columnas_vacias = ["Hs Extras", "Dep. Judicial", "Prepaga","Enlazados", "Ant. de Sueldos", "Adic. Ventas", "OBSERVACIONES"]
            
    # Leer el config.json desde la misma carpeta donde está el ejecutable (o .py si estás en desarrollo)
    if getattr(sys, 'frozen', False):
        BASE_PATH = os.path.dirname(sys.executable)  # cuando es .exe
    else:
        BASE_PATH = os.path.dirname(os.path.abspath(__file__))  # cuando es .py

    RUTA_CONFIG = os.path.join(BASE_PATH, 'config.json')

    with open(RUTA_CONFIG, 'r') as archivo:
        config_data = json.load(archivo)

    ruta_excel = config_data['salida_excel']

    