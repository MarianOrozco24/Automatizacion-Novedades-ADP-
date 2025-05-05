import pandas as pd
from datetime import datetime
import openpyxl
import mysql.connector.connection
import threading
import os 
from config import Config
import customtkinter as ctk
from tkinter import messagebox as mx
import sys, json
import traceback
from cryptography.fernet import Fernet
""" FUNCIONES """   
import subprocess


def desencriptar_credenciales ():
    # Cargar clave
    with open("variablesEntorno/clave_secreta.key", "rb") as archivo_clave:
        key = archivo_clave.read()

    fernet = Fernet(key)

    # Cargar credencial encriptada
    with open("credencial_encriptada.json", "rb") as archivo:
        credencial_encriptada = archivo.read()

    # Desencriptar
    credencial = fernet.decrypt(credencial_encriptada).decode()
    
    # Retornamos las credenciales
    return credencial



credenciales = desencriptar_credenciales()
# Ahora la podés usar, por ejemplo, para loguearte a una API o base de datos


def vpn_activa(ip_servidor):
    resultado = subprocess.run(["ping", "-n", "1", ip_servidor], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return resultado.returncode == 0


def conexion_db():
    try:
        cnx = mysql.connector.connect(host="172.16.20.15", user="user_aux", password=credenciales, database="db_omnia", use_pure=True)
        print("✅ Conexion a base de datos establecida correctamente")
        return cnx
    except ConnectionError:
        mx.showerror("Ha Ocurrido un error", "Ocurrio un error al intentar conectarse a la base de datos. Revise que el cable de ethernet este conectado correctamente. Si se conecta por VPN, revise que la misma este conectada")

    except Exception as error:
        traceback.print_exc()
        print(f"❌ Ocurrio un error al intentarse conectar a la base de datos\n \
            Error: {error}")
        
def query(cnx, query, tabla):
    try:
        cursor = cnx.cursor()
        cursor.execute(query)
        resultado = cursor.fetchall()
        if len(resultado) > 0:
            print(f"✅ Se extrajeron los datos correctamente. Cantidad de datos extraidos: {len(resultado)}")
        else:
            print(f"⚠️ La tabla {tabla} vino sin informacion")
            if tabla == "payroll" or tabla == "nomina":
                exit()
        # Cerrar cursor
        cursor.close()

        return resultado
    except Exception as error_query:
        mx.showerror("Error al consumir los datos", "Ha ocurrido un error al intentar consumir los datos. Contacte al administrador del programa")
        print(f"❌ Ocurrio un error al intentar ejecutar la consulta\n \
              Error: {error_query}")
        exit()


def cerrar_conexion_db (cnx):
    cnx.close()
    print("✅ Se cerro la conexion con base de datos")

""" CONSULTAS """
def consulta_payroll(month, year):
    payroll_extraction_query = f"""
    SELECT
    legajo,
    documento, 
    apellidos, 
    nombres, 
    equipo, 
    fecha, 
    codigo, 
    descripcion, 
    horas_programadas, 
    horas_trabajadas,
    usuario_neotel,
    usuario_avaya,
    campana, 
    sub_campana, 
    area, 
    puesto
    FROM payroll
    WHERE MONTH(fecha) = {month} AND YEAR(fecha) = {year} ;""" 
    return payroll_extraction_query

def consulta_nomina(month, year):
    nomina_extraction_query = f"""
    SELECT legajo, documento, apellidos, nombres 
    FROM asesores 
    where fecha_baja is null or (month(fecha_baja) >= {month} and year(fecha_baja) = {year})  ;
    """
    return nomina_extraction_query

def consulta_altas(month, year):
    nomina_extraction_query_fechas_altas = f"""
    SELECT legajo, fecha_alta FROM asesores WHERE month(fecha_alta) >= {month} and year(fecha_alta) = {year}"""

    return nomina_extraction_query_fechas_altas

def consulta_bajas(month, year):
    nomina_extraction_query_fechas = f"""
    select legajo, fecha_baja
    from asesores 
    where (month(fecha_baja) >= {month} and year(fecha_baja) = {year})"""

    return nomina_extraction_query_fechas

def procesamiento_de_datos(nomina, payroll, fechas_alta, fechas_baja, lista_columnas_vacias):
    try:
        ### Creacion de columnas en nomina que me realicen el conteo de dias dependiendo el n° de legajo
        nomina["Basico"] = 30

        # Contar cuántas veces aparece cada legajo con código "x, lPS, ART, LMED, " (D. Enfermedad)
        enfermedad = payroll[payroll['codigo'].isin(['CM','LPS',  'ART', 'LMED',])] \
                .groupby('legajo').size().reset_index(name='D. Enfermedad')

        # Contar cuántas veces aparece cada legajo con código "v" (Vacaciones)
        vacaciones = payroll[payroll['codigo'] == 'V'].groupby('legajo').size().reset_index(name='Vacaciones')

        # Contar cuántas veces aparece cada legajo con código "a" (Ausencias Injustificadas)
        feriado = payroll[payroll['codigo'] == 'DNL'].groupby('legajo').size().reset_index(name='Feriado')

        # Contar cuántas veces aparece cada legajo con código "L" (Ausencias Injustificadas) ||-> Enfermedades
        licencias = payroll[payroll['codigo'].isin(['LESP',  'MTM',  'CMh', 'LM', 'LFF'])] \
                .groupby('legajo').size().reset_index(name='Licencia')

        # Contar cuántas veces aparece cada legajo con código "a" (Ausencias Injustificadas)
        mudanza = payroll[payroll['codigo'] == 'M'].groupby('legajo').size().reset_index(name='D. Mudanza')


        # Contar cuántas veces aparece cada legajo con código "a" (Ausencias Injustificadas)
        estudio = payroll[payroll['codigo'] == 'CE'].groupby('legajo').size().reset_index(name='D. Estudio')


        # Contar cuántas veces aparece cada legajo con código "a" (Ausencias Injustificadas)
        suspension = payroll[payroll['codigo'] == 'S'].groupby('legajo').size().reset_index(name='D. Suspension')


        # Contar cuántas veces aparece cada legajo con código "a" (Ausencias Injustificadas)
        tramite = payroll[payroll['codigo'] == 'DT'].groupby('legajo').size().reset_index(name='D. Tramite')


        # Contar cuántas veces aparece cada legajo con código "a" (Ausencias Injustificadas)
        falta_injustificada = payroll[payroll['codigo'] == 'X'].groupby('legajo').size().reset_index(name='Falta Inj.')



        # Combinar todos los conteos en un solo DataFrame
        conteos = enfermedad.merge(vacaciones, on='legajo', how='outer') \
                            .merge(feriado, on='legajo', how='outer') \
                            .merge(licencias, on='legajo', how='outer') \
                            .merge(mudanza, on='legajo', how='outer') \
                            .merge(estudio, on='legajo', how='outer') \
                            .merge(suspension, on='legajo', how='outer') \
                            .merge(tramite, on='legajo', how='outer') \
                            .merge(falta_injustificada, on='legajo', how='outer') 

        nomina = nomina.merge(conteos, on='legajo', how='left')

        # Completamos las columnas que se llenan a mano
        for col in lista_columnas_vacias:
            nomina[col] = None

        # Columna observaciones
        nomina['fecha_baja'] = nomina['legajo'].map(fechas_baja.set_index('legajo')['fecha_baja'])
        nomina['fecha_alta'] = nomina['legajo'].map(fechas_alta.set_index('legajo')['fecha_alta'])

        # nomina = nomina.merge(fechas_baja[['legajo', 'fecha_baja', 'fecha_alta']], on='legajo', how='left')

        # Asegurar que la columna fecha_baja es de tipo datetime
        nomina['fecha_baja'] = pd.to_datetime(nomina['fecha_baja'], errors='coerce')

        # Extraer el día como número entero, manejando NaN correctamente
        nomina['dia_baja'] = nomina['fecha_baja'].dt.day.fillna(0).astype(int)

        nomina["dia_baja"] = nomina["dia_baja"].clip(upper=30)

        # Restar solo cuando haya una fecha de baja registrada
        nomina['Basico'] = nomina['Basico'] - nomina['dia_baja'].where(nomina['fecha_baja'].notna(), 0)


        # Ajustamos los numeros de la columna Basico, ya que sino solamente nos queda el restante de dias y no
        # los facturables
        # # Asegurar que Basico no sea menor a 30
        nomina['Basico'] = nomina['Basico'].where(nomina['Basico'] >= 30, 30 - nomina['Basico'])

        # Convertimos la columna fecha_alta a datetime
        nomina['fecha_alta'] = pd.to_datetime(nomina['fecha_alta'], errors='coerce')

        # Extraer el día como número entero, manejando NaN correctamente
        nomina['dia_alta'] = nomina['fecha_alta'].dt.day.fillna(0).astype(int)

        nomina["dia_alta"] = nomina["dia_alta"].clip(upper=30)

        nomina['Basico'] = nomina['Basico'] - nomina['dia_alta'].where(nomina['fecha_alta'].notna(), 0)



        # Sumar todas las columnas que afectan al cálculo
        descuento = nomina[['Falta Inj.', 'D. Mudanza', 'D. Estudio', 'D. Enfermedad', 'Licencia', 'D. Tramite']].sum(axis=1)

        # Aplicar descuento solo cuando 'Falta Inj.' > 0, dejando el resto sin cambios
        nomina['Basico'] = nomina['Basico'] - descuento.where(nomina['Falta Inj.'] > 0, 0)

        # Ajustar 'Basico' solo si 'Falta Inj.' es nulo y hay una observación
        nomina.loc[(nomina['Falta Inj.'].isna()) & (nomina["fecha_baja"].notna()), "Basico"] = nomina['dia_baja']

        """ renombramos las columnas alta/ baja"""
        nomina = nomina.rename(columns={"fecha_baja" : "Fecha baja",
                                         "fecha_alta" : "Fecha alta", 
                                         "legajo" : "Legajo", 
                                         "documento" : "Documento", 
                                         "apellidos" : "Apellido", 
                                         "nombres" :  "Nombre"})

        return nomina
    except Exception:
        mx.showerror("Error","Ha ocurrido un error, contacte al administrador del programa")
        print(f"❌ Ocurrio un error\n\
              Error: {Exception}")


def main(month,year):
    month = int(month.get())
    year = int(year.get())

    date_obj = datetime(year, month, 1)  
    date_str = date_obj.strftime("%m/%Y")

    name_month = date_obj.strftime("%B")
    file_name = f"{name_month} {year}.xlsx"
    file_name


    """ EXTRACCION DE INFORMACION """
    if not vpn_activa("172.16.20.15"):
        print("❌ No estás conectado a la VPN. Conectate antes de ejecutar el programa.")
        mx.showerror("Error de conexion", "No se obtuvo informacion de la base de datos. En caso de usar VPN, revise la conexion y si esta conectada, presione reconectar e intente nuevamente")
        exit()

    
    # Columnas payroll
    columnas_payroll = ["legajo","documento", "apellidos", "nombres","equipo","fecha", "codigo", "descripcion", 
    "horas_programadas","horas_trabajadas","usuario_neotel","usuario_avaya","campana", "sub_campana","area", "puesto"]
    # Consumimos payroll
    payroll_extraction_query = consulta_payroll(month, year)
    cnx = conexion_db()
    resultado = query(cnx,payroll_extraction_query, "payroll")
    payroll = pd.DataFrame(resultado, columns=columnas_payroll)
    
    
    # Columnas nomina
    columnas_nomina = ["legajo","documento", "apellidos", "nombres"]
    # Consulta nomina
    nomina_extraction_query = consulta_nomina(month, year)
    resultado = query(cnx,nomina_extraction_query, "nomina")
    nomina = pd.DataFrame(resultado, columns=columnas_nomina)
    
    
    # Consultas bajas
    nomina_extraction_query_fechas_bajas = consulta_bajas(month, year)
    resultado = query(cnx,nomina_extraction_query_fechas_bajas, "bajas")
    fechas_baja = pd.DataFrame(resultado, columns=["legajo", "fecha_baja"])


    # Consultas altas
    nomina_extraction_query_fechas_altas = consulta_altas(month, year)
    resultado = query(cnx,nomina_extraction_query_fechas_altas, "bajas")
    fechas_altas = pd.DataFrame(resultado, columns=["legajo", "fecha_alta"])
    
    # Cerrar conexion a base de datos
    cerrar_conexion_db(cnx)

    
    """ PROCESAMIENTO DE DATOS """
    df_final = procesamiento_de_datos(nomina, payroll, fechas_altas, fechas_baja, Config.lista_columnas_vacias)

    """ LIMPIEZA DF FINAL"""
    # Todos los valores de la columnas basico que sean = 30 los vamos a colocar en null
    df_final.loc[df_final['Basico'] == 30, 'Basico'] = None
    
    # Eliminamos las columnas extras
    df_final = df_final.drop(axis=1, columns=[ "dia_baja", "dia_alta"])


    """ Extraccion de resultado """  

    # exportamos a excel para analizar 
    try:
        df_final = df_final.sort_values(by=['Apellido', 'Nombre'], ascending=[True, True])
        df_final.to_excel(f"{Config.ruta_excel}Novedades {file_name}", index=False)
        print(f"✅ Se extrajo correctamente el archivo {file_name} en {Config.ruta_excel}")
        mx.showinfo("Exportacion exitosa", f"Se exporto correctamente el archivo Novedades - {file_name} en {Config.ruta_excel}")
    except PermissionError as error_permisos:
        print(f"❌ Ocurrio un error al exportar el archivo\n\
            Si hay un archivo con el mismo nombre del que se va a exportar, en este caso\n\
            {file_name} cierrelo antes de ejecutar el programa")
        
        mx.showerror("Ocurrio un error al exportar el archivo", f"Si hay un archivo con el mismo nombre del que se va a exportar, en este caso {file_name} cierrelo antes de ejecutar el programa")
    except Exception as error_exportacion:
        print(f"❌ Ocurrio un error al exportar el archivo\n \
            Error: {error_exportacion}")
        mx.showerror("Ocurrio un error al exportar el archivo", f"Ocurrio un error al exportar el archivo {file_name}. Contacte al administrador del programa")



""" VENTANA PRINCIPAL """
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")
ventana_principal = ctk.CTk()
ventana_principal.title("Ventana Principal")
ventana_principal.geometry("700x400")

# titulo
label_title = ctk.CTkLabel(ventana_principal, text="Automatizacion Novedades", font=("helvetica", 24, "bold"))
label_title.pack(pady=35, padx=5)

# frame_mes 
frame_mes = ctk.CTkFrame(ventana_principal, fg_color="#2C2F33")
frame_mes.pack(pady=15, padx=5)
entry_mes = ctk.CTkEntry(frame_mes, placeholder_text="Mes (Ej: 5)",fg_color="#2C2F33" )
entry_mes.pack(pady=1, padx=1)
# frame_año 
frame_año = ctk.CTkFrame(ventana_principal, fg_color="#2C2F33")
frame_año.pack(pady=15, padx=5)
entry_año = ctk.CTkEntry(frame_año, placeholder_text="Año (Ej: 2025)", fg_color="#2C2F33")
entry_año.pack(pady=2, padx=2)

frame_button = ctk.CTkFrame(ventana_principal, fg_color="#6A0DAD")
frame_button.pack(pady=60, padx=5)
button_continue = ctk.CTkButton(frame_button, text="Ejecutar",fg_color="#2C2F33", command=lambda: main(entry_mes, entry_año))
button_continue.pack(pady=2, padx=2)

ventana_principal.bind("<Return>", lambda event: main(entry_mes, entry_año))



ventana_principal.mainloop()


