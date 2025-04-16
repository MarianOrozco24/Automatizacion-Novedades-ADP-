# Automatizacion-Adm-Personal

Objetivo
Este archivo automatiza la generación de un reporte mensual en Excel a partir de datos almacenados en una base de datos MySQL. El reporte contiene información de nómina y asistencia de empleados, organizada por días de trabajo, licencias, ausencias y otros eventos relevantes.

Proceso General
Configuración: Define el mes y año para el reporte (ejemplo: Marzo 2025).

Conexión a la Base de Datos: Extrae datos de dos tablas:

payroll (registros diarios de asistencia).

asesores (datos básicos de empleados).

Procesamiento:

Calcula días específicos (vacaciones, licencias, ausencias, etc.) para cada empleado.

Combina los datos en una estructura lista para exportar.

Generación del Reporte: Guarda la información en un archivo Excel con el nombre [Mes] [Año].xlsx (ejemplo: March 2025.xlsx).

Componentes Clave
1. Configuración Inicial
Mes y Año: Se definen manualmente (ejemplo: month = 3, year = 2025).

Nombre del Archivo: Se genera automáticamente (ejemplo: March 2025.xlsx).

2. Conexión a la Base de Datos
Credenciales:

Servidor: 172.16.20.15

Usuario: mariano_orozco

Base de datos: db_omnia

Tablas Utilizadas:

payroll: Registros diarios de horas trabajadas, códigos de eventos (vacaciones, licencias, etc.).

asesores: Lista de empleados con datos básicos (legajo, nombre, apellidos).

3. Procesamiento de Datos
Datos de Asistencia (payroll):

Se filtran registros del mes y año configurados.

Se cuentan días por tipo de evento (ejemplo: días de enfermedad, vacaciones, feriados).

Datos de Empleados (asesores):

Se combinan con los conteos de días para crear una tabla consolidada.

Se agregan columnas vacías para información adicional (ejemplo: horas extras, observaciones).

4. Salida Final
Archivo Excel:

Ordenado alfabéticamente por apellidos y nombres.

Incluye todas las columnas relevantes para la nómina.

Beneficios
Automatización: Elimina la necesidad de cálculos manuales y extracción repetitiva de datos.

Precisión: Reduce errores humanos al estandarizar el proceso.

Eficiencia: Genera el reporte en segundos, listo para revisión o distribución.

Requisitos Técnicos
Dependencias:

Python 3.x

Bibliotecas: pandas, mysql.connector, openpyxl.

Acceso a la Base de Datos: Credenciales válidas y conexión al servidor 172.16.20.15.

Pasos para Ejecutar
Asegurar que las credenciales de la base de datos sean correctas.

Definir el mes y año deseados en la sección de configuración.

Ejecutar el cuaderno (app.ipynb).

El archivo Excel se generará automáticamente.



Comando ejecutable:
pyinstaller --onefile ^
 --icon=icono/icon.ico ^
 --hidden-import=mysql ^
 --hidden-import=mysql.connector ^
 --hidden-import=mysql.connector.plugins.mysql_native_password ^
 --add-data "config.json;." ^
 --add-data "C:/Users/Usuario/Desktop/Codigos Planificacion Estrategica/Automatizacion-Adm-Personal/.venv/Lib/site-packages/mysql/connector/locales;mysql/connector/locales" ^
 app.py


