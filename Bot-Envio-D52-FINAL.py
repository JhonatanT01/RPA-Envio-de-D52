import subprocess
import pyautogui
import time
import win32com.client
import pyperclip
import webbrowser
from datetime import datetime 
import tkinter as tk
from tkinter import messagebox


# Ruta de los scripts VBScript
ruta_EnvioD52 = r"C:/Users/NCACUAXSERLO/AppData/Roaming/SAP/SAP GUI/Scripts/ReporteD52R.vbs"

def abrir_url(url, espera=10):
    webbrowser.open(url)
    time.sleep(espera)  # Esperar a que la página cargue completamente
    
    # Esperar a que la ventana de Google Sheets esté activa
    ventana_abierta = False
    intentos = 0
    
    while not ventana_abierta and intentos < 5:  # Máximo 5 intentos
        ventanas = pyautogui.getWindowsWithTitle("Google Sheets")
        if ventanas:
            ventanas[0].activate()  # Activar la primera ventana que coincida
            ventana_abierta = True
        else:
            time.sleep(2)
            intentos += 1


# Función para ejecutar scripts VBScript
def ejecutar_script_vbscript(ruta_script):
    try:
        proceso = subprocess.Popen(['cscript.exe', ruta_script], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        time.sleep(2)  # Ajustar según sea necesario
        pyautogui.press('ENTER')
        stdout, stderr = proceso.communicate()
        salida = stdout.decode('utf-8', errors='ignore')
        print("Salida del script:", salida)
        
        if stderr:
            errores = stderr.decode('utf-8', errors='ignore')
            print("Errores:", errores)
    except Exception as e:
        print(f"Ocurrió un error: {e}")

# Ejecutar script MB52
ejecutar_script_vbscript(ruta_EnvioD52)
time.sleep(2)
# Selecionar  de del Reporte de D52
excel = win32com.client.Dispatch("Excel.Application")
workbook_origen = excel.ActiveWorkbook
worksheet_origen = workbook_origen.ActiveSheet
worksheet_origen.Range("A2:S400").Select()
time.sleep(2)
excel.Selection.Copy()


time.sleep(3)
#Tabal de D52
url130 = 'https://docs.google.com/spreadsheets/d/1E__HUReE_pQI6EFK9GVaOwnEReePVPCdOX5sK-1RWDM/edit?pli=1&gid=0#gid=0'
subprocess.run(["cmd", "/c", "start", "chrome", "--new-window", url130])
time.sleep(10)
pyautogui.hotkey('ctrl', 'shift', 'alt', '9')
pyautogui.press('down')
time.sleep(1)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')
time.sleep(1)
pyautogui.hotkey('ctrl', 'V')
time.sleep(2)



for _ in range(20):
    pyautogui.press('right')
    time.sleep(0.1) 

#ejecutara macro
pyautogui.hotkey('ctrl', 'shift', 'alt', '7')

time.sleep(3)
pyautogui.hotkey('ctrl', 'c')
time.sleep(6)

# Variable de fecha actual

fecha_actual = datetime.now().strftime('%d-%B')

# Abrir Gmail en la ventana de redacción de un nuevo correo
gmail_url99 = 'https://mail.google.com/mail/u/0/?pli=1#inbox?compose=new'
webbrowser.open(gmail_url99)
time.sleep(11)  # Esperar a que cargue Gmail

# Pegar contenido y borrar caracteres innecesarios
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)


# BORRAR CARACTERES ESPECIALES 

# Obtener el contenido del portapapeles
contenido = pyperclip.paste()
# Lista de palabras a eliminar
palabras_a_eliminar = [ "CORREO","COREOSCOREO","COREOS"]
# Separar por líneas y filtrar correos válidos
correos = contenido.split("\n")
correos_limpios = [correo.strip() for correo in correos if correo.strip() and correo.strip() not in palabras_a_eliminar]
# Volver a unir los correos limpios
contenido_limpio = "\n".join(correos_limpios)

# Copiar de nuevo al portapapeles solo los correos válidos
pyperclip.copy(contenido_limpio)

# Pegar los correos limpios en el campo de destinatarios
pyautogui.hotkey('ctrl', 'v')

#FIN 
time.sleep(2)
###ENVIO DE CORREO##S   
# Activar el campo CC y agregar destinatarios
pyautogui.hotkey('ctrl', 'shift', 'c')
time.sleep(4)  # Esperar 2 segundos
cc_destinatarios = ['July Esperanza Cuellar Delgadillo', 'Yesid Sanchez']
# Escribir el primer correo
pyautogui.write(cc_destinatarios[0])
time.sleep(6)  # Esperar 2 segundos
# Presionar tab para avanzar al siguiente campo
pyautogui.press('tab')
# Escribir el segundo correo
time.sleep(5)  # Esperar 3 segundos antes de escribir el segundo
pyautogui.write(cc_destinatarios[1])
time.sleep(5)  # Esperar 2 segundos
pyautogui.press('tab')
pyautogui.press('tab')
time.sleep(5)
# Escribir el asunto del correo
asunto = 'Clientes más Representativos del D52'
pyperclip.copy(asunto)
time.sleep(1)
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)
pyautogui.press('tab')
time.sleep(6)

# Variable de fecha actual
fecha_actual = datetime.now().strftime('%d-%m-%Y')
# Escribir el cuerpo del correo
mensaje = f'''
Cordial saludo,

Envío clientes más representativos en Devoluciones en buen estado del {fecha_actual}.

D52:
'''
pyperclip.copy(mensaje)
time.sleep(1)
pyautogui.hotkey('ctrl', 'v')
time.sleep(5)




#ABRIR URL DE D52 TABLA 
url178 = "https://docs.google.com/spreadsheets/d/1uUrB_Ia5oLx9JUG_D5apsWB6Xy8rLeuO/edit?gid=2023613478#gid=2023613478"
subprocess.run(["cmd", "/c", "start", "chrome", "--new-window", url178])  # Abre otra nueva ventana
#Copir Datos de Tabla de reporte
time.sleep(4)
pyautogui.press('down')
time.sleep(4)  
pyautogui.hotkey('ctrl', 'a')
time.sleep(2)
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt',  'tab') 
time.sleep(5)
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)

for _ in range(4):
    pyautogui.press("down")
    time.sleep(0.2)  # Pequeña pausa entre cada pulsación


time.sleep(1)
pyperclip.copy("RPA D52R-Devoluciones")
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)
time.sleep(1)
pyautogui.press('tab')  
time.sleep(1)


# Crear la ventana principal oculta
root = tk.Tk()
root.withdraw()  # Ocultar la ventana principal

# Mostrar ventana emergente
messagebox.showinfo("Proceso finalizado", "¡Ejecución completada!\nTodo ha terminado.")

# Cerrar la aplicación después de mostrar el mensaje
root.destroy()