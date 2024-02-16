import re
from datetime import date, timedelta, datetime
from openpyxl import Workbook, load_workbook
from colorama import Fore
from bs4 import BeautifulSoup
import requests 

url_datos = "https://backend.jaha.com.py/compensaciones/tarjetas_diferencias_eventos"


##########################################################################################

# url_ingresos_ps = "https://backoffice.jaha.com.py/reportes/ingresospsdetallado"

# client = requests.Session()

# html = client.get('https://backoffice.jaha.com.py/reportes/ingresospsdetallado').content
# soup = BeautifulSoup(html, 'html.parser')

# print(soup.find(name="csrf-token").text)




##########################################################################################
################################### OBTENCION DE DATOS ###################################
##########################################################################################

result = requests.get(url_datos)
content = result.text

fecha_referencia = date.today().strftime('%Y-%m-%d')

posicion_inicio = content.find(fecha_referencia)
datos_ultima_compensacion = content[posicion_inicio:].replace("</pre>", "")

# Iniciamos un indice para barrido de fechas para obtener la última compensación disponible
i = 1

# Bucle que obtiene fecha y datos de la compensacion mas reciente
while datos_ultima_compensacion == ">": # Se compara con valor de datos_ultima_compensacion cuando no hubo compensaciones en la fecha de hoy
    fecha_compensacion = date.today() - timedelta(days=i)
    fecha_compensacion = fecha_compensacion.strftime('%Y-%m-%d')
    posicion_inicio = content.find(fecha_compensacion)
    datos_ultima_compensacion = content[posicion_inicio:].replace("</pre>", "")
    i = i + 1

print(datos_ultima_compensacion)

##########################################################################################
###################################### MODULO EXCEL ######################################
##########################################################################################

# excelname = datetime.today()

#excelname = excelname.strftime('%Y-%m-%d %H_%M_%S.xlsx')
#wb = Workbook()
# ws = wb.create_sheet('Hojarda')

excelName = "Planilla de Tiempos de Compensaciones de PY.xlsx"
hojaName = "Datos Backend"
excelObject = load_workbook(excelName)


hojaObject = excelObject[hojaName]

value = hojaObject['B21'].value #5674

print(value)

fecha_index = None

# Ciclo para obtener el indice de la celda donde esta la ultima compensacion
index_tabla = 7 # Un valor de referencia para reducir la cantidad de iteraciones
while fecha_index != fecha_compensacion:
    index_tabla = index_tabla + 1
    fecha_index = hojaObject['A{index_tabla}'.format(index_tabla = index_tabla)].value.strftime('%Y-%m-%d')

respaldo_compensacion = datos_ultima_compensacion

# Sustituye espacios multiples por puntos y comas
datos_ultima_compensacion = re.sub(r"\s+", ";", datos_ultima_compensacion)

print(datos_ultima_compensacion)


hojaObject['B{index_tabla}'.format(index_tabla = index_tabla)] = datos_ultima_compensacion
        
# print("Posicion " + str(index_tabla))
# print(fecha_index)


# Ejemplo para llenar celda
# hojaObject['B22'] = "PUTO EL QUE LEE"




excelObject.save(excelName)



#wb.save(excelname)



# ################# Variables #################
# Tarjetas_Excluidas_Acumuladas	
# Tarjetas_en_Espera_Acumuladas	
# Tarjetas_Compensadas	
# Eventos_Compensados	
# Pago_TDP_a_EPAS_(GS)	
# Pago_TDP_a_EPAS_(USD)	
# Eventos_Excluidos_Cantidad
# Eventos_Excluidos_Porcent
# Eventos_En_Espera_Cantidad
# Eventos_En_Espera_Porcent
# Eventos_Listo_Cantidad
# Eventos_Listo_Porcent
# Eventos_en_Error_Cantidad
# Eventos_en_Error_Porcent
# Eventos_Perdidos_Cantidad
# Eventos_Perdidos_Porcent
# Eventos_Totales_Cantidad
# Eventos_Totales_Porcent

#print("\n" + result2.text)