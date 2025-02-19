from google import procesar_lote_google
from arcGis import procesar_lote_arcgis
from predictiveSearchQAar import procesar_lote_predictiveAR
from predictiveSearchQAcl import procesar_lote_predictiveCL
from locationEsriBase import procesar_lote_locationEsri
from locationHereBase import procesar_lote_locationHere
from locationEsriEnrich import procesar_lote_locationEsriEnrich
from locationHereEnrich import procesar_lote_locationHereEnrich
from addressPEProd import procesarLoteAddressPEProd
from phonePEQA import procesarLotePhonePeQA
from geotypes import *
from comunes import *

BLACK = "\033[30m"
RED = "\033[31m"
GREEN = "\033[32m"
YELLOW = "\033[33m"
RED = "\033[34m"
MAGENTA = "\033[35m"
CYAN = "\033[36m"
WHITE = "\033[37m"
RESET = "\033[0m"
BOLD = "\033[1m"
UNDERLINE = "\033[4m"
BACKGROUND_RED = "\033[41m"
BACKGROUND_GREEN = "\033[42m"

print(CYAN + BOLD + "Merlin Data Quality 2024\n" + RESET)

print(MAGENTA + BOLD + ".: BIENVENID@ AL PROCESADOR DE LOTES DE LIBRERIAS VERSIÓN 11 :.\n" + RESET)
print("Fecha de versión: 29/08/2024\n")
print("OPCIÓN 1 - Procesar el lote por Google")
print("OPCIÓN 2 - Procesar el lote por ArcGis")
print("OPCIÓN 30 - Procesar el lote por Merlin Predictive Search AR (QA)")
print("OPCIÓN 31 - Procesar el lote por Merlin Predictive Search CL (QA)")
print("OPCIÓN 40 - Procesar por AWS Location Services en ESRI")
print("OPCIÓN 41 - Procesar por AWS Location Services en HERE")
print("OPCIÓN 5 - Procesar por Merlin Address Perú (Prod)")
print("OPCIÓN 6 - Procesar por Merlin Phone Perú (QAs)")

opcion = str(input(YELLOW + "Ingresá el número de opción acá --> " + RESET))

if opcion == "1":
    procesar_lote_google('./Lotes a procesar/entrada_google.xlsx')
    mensaje_saludo()
    input(YELLOW + "Presione cualquier tecla para cerrar" + RESET) 
elif opcion == "2":
    procesar_lote_arcgis()
    mensaje_saludo()
    input(YELLOW + "Presione cualquier tecla para cerrar" + RESET) 
elif opcion == '30':
    procesar_lote_predictiveAR()
    mensaje_saludo()
    input(YELLOW + "Presione cualquier tecla para cerrar" + RESET)
elif opcion == '31':
    procesar_lote_predictiveCL()
    mensaje_saludo()
    input(YELLOW + "Presione cualquier tecla para cerrar" + RESET)
elif opcion == '40':
    procesar_lote_locationEsriEnrich()
    mensaje_saludo()
    input(YELLOW + "Presione cualquier tecla para cerrar" + RESET)
elif opcion == '41':
    procesar_lote_locationHereEnrich()
    mensaje_saludo()    
    input(YELLOW + "Presione cualquier tecla para cerrar" + RESET)
elif opcion == '6':
    procesarLotePhonePeQA()
    mensaje_saludo()
    input(YELLOW + "Presione cualquier tecla para cerrar" + RESET)  
elif opcion == '5':
    procesarLoteAddressPEProd()
    mensaje_saludo()
    input(YELLOW + "Presione cualquier tecla para cerrar" + RESET)
else:
    mensaje_saludo()    
    input(YELLOW + "Presione cualquier tecla para cerrar" + RESET) 

