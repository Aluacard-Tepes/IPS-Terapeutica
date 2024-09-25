import pandas as pd
from jinja2 import Environment, FileSystemLoader
import datetime
import pdfkit
import os
import locale
from tkinter import filedialog
from tkinter import Tk

# Función para abrir el cuadro de diálogo de selección de archivo
def seleccionar_archivo():
    root = Tk()
    root.withdraw()  # Oculta la ventana principal de tkinter
    archivo_seleccionado = filedialog.askopenfilename(title="Seleccione la base de datos (Excel)", filetypes=[("Archivos de Excel", "*.xlsx")])
    return archivo_seleccionado

# Seleccionar la base de datos
ruta_base_datos = seleccionar_archivo()

# Verificar que se haya seleccionado un archivo
if not ruta_base_datos:
    print("No se seleccionó ningún archivo. Saliendo del programa.")
    exit()

# Función para abrir el cuadro de diálogo de selección de carpeta
def seleccionar_carpeta():
    root = Tk()
    root.withdraw()  # Oculta la ventana principal de tkinter
    carpeta_destino = filedialog.askdirectory(title="Seleccione la carpeta de destino")
    return carpeta_destino

# BASE IPS ESTO CAMBIA POR MES
BDIps = pd.read_excel(ruta_base_datos)

# Configura la ruta al ejecutable de wkhtmltopdf
config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')  # Reemplaza con la ruta correcta

# Obtener el número total de filas en el DataFrame
num_filas = len(BDIps)

# VARIABLES FECHA
locale.setlocale(locale.LC_TIME,'es_ES.UTF-8')
fecha_actual = datetime.datetime.now()
nombre_mes_actual = fecha_actual.strftime('%B')
VarMes = nombre_mes_actual
VarAño = datetime.datetime.now().year

carpeta_destino = seleccionar_carpeta()


for i in range(num_filas):
    fila = BDIps.iloc[i]  # Obtener la fila actual

    # Variables de la fila actual
    VarID = fila['ID']
    VarTipID = fila['Tipo de Documento']
    VarName = fila['Nombre']
    
    # Variables Sesiones TO
    VarTOSm1 = fila['T.O Semana 1']
    VarTOSm2 = fila['T.O Semana 2']
    VarTOSm3 = fila['T.O Semana 3']
    VarTOSm4 = fila['T.O Semana 4']
    VarTOSm5 = fila['T.O Semana 5']
    VarTOSm6 = fila['T.O Semana 6']
    VarTOSuma = fila['Suma T.O']
    
    # Variables Sesiones Fono
    VarFOSm1 = fila['FONO Semana 1']
    VarFOSm2 = fila['FONO Semana 2']
    VarFOSm3 = fila['FONO Semana 3']
    VarFOSm4 = fila['FONO Semana 4']
    VarFOSm5 = fila['FONO Semana 5']
    VarFOSm6 = fila['FONO Semana 6']
    VarFOSuma = fila['Suma FONO']
    
    # Variables Sesiones Psico
    VarPSSm1 = fila['PSICO Semana 1']
    VarPSSm2 = fila['PSICO Semana 2']
    VarPSSm3 = fila['PSICO Semana 3']
    VarPSSm4 = fila['PSICO Semana 4']
    VarPSSm5 = fila['PSICO Semana 5']
    VarPSSm6 = fila['PSICO Semana 6']
    VarPSSuma = fila['Suma Psico']
    
    # Variables Sesiones Fisio
    VarFISm1 = fila['FISIO Semana 1']
    VarFISm2 = fila['FISIO Semana 2']
    VarFISm3 = fila['FISIO Semana 3']
    VarFISm4 = fila['FISIO Semana 4']
    VarFISm5 = fila['FISIO Semana 5']
    VarFISm6 = fila['FISIO Semana 6']
    VarFISuma = fila['Suma Fisio']
    
    # Variables Sesiones Rehabilitacion cognitiva
    VarRHSm1 = fila['RHC Semana 1']
    VarRHSm2 = fila['RHC Semana 2']
    VarRHSm3 = fila['RHC Semana 3']
    VarRHSm4 = fila['RHC Semana 4']
    VarRHSm5 = fila['RHC Semana 5']
    VarRHSm6 = fila['RHC Semana 6']
    VarRHSuma = fila['Suma RHC']
    
    # Variables Sesiones PSICOLOGIA CLINICA
    VarCLSm1 = fila['CLI Semana 1']
    VarCLSm2 = fila['CLI Semana 2']
    VarCLSm3 = fila['CLI Semana 3']
    VarCLSm4 = fila['CLI Semana 4']
    VarCLSm5 = fila['CLI Semana 5']
    VarCLSm6 = fila['CLI Semana 6']
    VarCLSuma = fila['Suma CLI']   
     
    # Variables Sesiones TO GRUPAL
    VarTGSm1 = fila['TGR Semana 1']
    VarTGSm2 = fila['TGR Semana 2']
    VarTGSm3 = fila['TGR Semana 3']
    VarTGSm4 = fila['TGR Semana 4']
    VarTGSm5 = fila['TGR Semana 5']
    VarTGSm6 = fila['TGR Semana 6']
    VarTGSuma = fila['Suma TGR']
    
    # Variable Suma Total Sesiones
    varSumT = fila['Suma Total']
    varLogo = 'IPSLOGO.PNG'

    Env = Environment(loader=FileSystemLoader('A:/IPS SCRIPS/Desarrollo Scrip Facturación'))
    template = Env.get_template('FACTURA.HTML')

    Usuario = {
            'VarID': VarID, 
            'VarTipoID': VarTipID, 
            'VarName': VarName, 
            'VarMes': VarMes, 
            'VarAño': VarAño,
            'VarTOSm1': VarTOSm1,
            'VarTOSm2': VarTOSm2,
            'VarTOSm3': VarTOSm3,
            'VarTOSm4': VarTOSm4,
            'VarTOSm5': VarTOSm5,
            'VarTOSm6': VarTOSm6,
            'VarTOSuma': VarTOSuma,
            'VarFOSm1': VarFOSm1,
            'VarFOSm2': VarFOSm2,
            'VarFOSm3': VarFOSm3,
            'VarFOSm4': VarFOSm4,
            'VarFOSm5': VarFOSm5,
            'VarFOSm6': VarFOSm6,
            'VarFOSuma': VarFOSuma,
            'VarPSSm1': VarPSSm1,
            'VarPSSm2': VarPSSm2,
            'VarPSSm3': VarPSSm3,
            'VarPSSm4': VarPSSm4,
            'VarPSSm5': VarPSSm5,
            'VarPSSm6': VarPSSm6,
            'VarPSSuma': VarPSSuma,
            'VarFISm1': VarFISm1,
            'VarFISm2': VarFISm2,
            'VarFISm3': VarFISm3,
            'VarFISm4': VarFISm4,
            'VarFISm5': VarFISm5,
            'VarFISm6': VarFISm6,
            'VarFISuma': VarFISuma,
            'VarRHSm1': VarRHSm1,
            'VarRHSm2': VarRHSm2,
            'VarRHSm3': VarRHSm3,
            'VarRHSm4': VarRHSm4,
            'VarRHSm5': VarRHSm5,
            'VarRHSm6': VarRHSm6,
            'VarRHSuma': VarRHSuma,
            'VarCLSm1': VarCLSm1,
            'VarCLSm2': VarCLSm2,
            'VarCLSm3': VarCLSm3,
            'VarCLSm4': VarCLSm4,
            'VarCLSm5': VarCLSm5,
            'VarCLSm6': VarCLSm6,
            'VarCLSuma': VarCLSuma,
            'VarTGSm1': VarTGSm1,
            'VarTGSm2': VarTGSm2,
            'VarTGSm3': VarTGSm3,
            'VarTGSm4': VarTGSm4,
            'VarTGSm5': VarTGSm5,
            'VarTGSm6': VarTGSm6,
            'VarTGSuma': VarTGSuma,
            'varSumT': varSumT,
             }

    html = template.render(Usuario)
    
    # Crear un archivo HTML único para cada fila
    html_file_name = os.path.join(carpeta_destino,f'nuevohtml_{VarName}_{VarID}.html')
    with open(html_file_name, 'w',encoding='utf-8') as f:
        f.write(html)
    
    options = {
        'user-style-sheet': 'FACTURA.CSS',
        'encoding': 'UTF-8',
    }
    
    # Generar un PDF único para cada fila
    pdf_file_name = os.path.join(carpeta_destino, f'{VarName}_{VarID}.pdf')
    pdfkit.from_file(html_file_name, pdf_file_name, configuration=config, options=options)
    os.remove(html_file_name)
    
    print(f'Procesado: {VarName}_{VarID}.pdf')
    print(f'Faltan: {num_filas - i - 1} archivos por procesar')