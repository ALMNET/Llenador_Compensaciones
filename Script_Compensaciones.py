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

excelName = datetime.today()
excelName = excelName.strftime('%Y-%m-%d %H_%M_%S.xlsx')
excelObject = Workbook()
excelObject.save(excelName)
hojaName = 'Hojarda'
hojaObject = excelObject.create_sheet(hojaName)
#hojaObject = excelObject[hojaName]
excelObject.save(excelName)


# excelName = "Planilla de Tiempos de Compensaciones de PY.xlsx"
# hojaName = "Datos Backend"
# excelObject = load_workbook(excelName)


# hojaObject = excelObject[hojaName]

# value = hojaObject['B21'].value #5674

# print(value)

fecha_index = None

# Ciclo para obtener el indice de la celda donde esta la ultima compensacion
# index_tabla = 7 # Un valor de referencia para reducir la cantidad de iteraciones
# while fecha_index != fecha_compensacion:
#     index_tabla = index_tabla + 1
#     fecha_index = hojaObject['A{index_tabla}'.format(index_tabla = index_tabla)].value.strftime('%Y-%m-%d')

# respaldo_compensacion = datos_ultima_compensacion

# Sustituye espacios multiples por puntos y comas
datos_ultima_compensacion = re.sub(r"\s+", ";", datos_ultima_compensacion)

print(datos_ultima_compensacion)


# hojaObject['A1'] = datos_ultima_compensacion

# excelObject.save(excelName)


# hojaObject['B{index_tabla}'.format(index_tabla = index_tabla)] = datos_ultima_compensacion
        
# print("Posicion " + str(index_tabla))
# print(fecha_index)


# Ejemplo para llenar celda
# hojaObject['B22'] = "PUTO EL QUE LEE"




# excelObject.save(excelName)



#wb.save(excelname)



# ################# Variables #################
# extracted_text = datos_ultima_compensacion[:datos_ultima_compensacion.find(";")]
# datos_ultima_compensacion = datos_ultima_compensacion[datos_ultima_compensacion.find(";") + 1:]

def extractor(texto, fila, columna):
    textoExtraido = texto[:texto.find(";")]
    texto = texto[texto.find(";") + 1:]

    hojaObject['{fila}{columna}'.format(fila = fila, columna = columna)] = textoExtraido
 
    return textoExtraido, texto

fila_index = 65 # 'A'

Fecha_base, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 0), '1')
Tarjetas_Excluidas_Acumuladas, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 1), '1')
Tarjetas_en_Espera_Acumuladas, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 2), '1')
Tarjetas_Compensadas, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 3), '1')
Eventos_Compensados, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 4), '1')
Pago_TDP_a_EPAS_GS, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 5), '1')
Pago_TDP_a_EPAS_USD, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 6), '1')
Eventos_Excluidos_Cantidad, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 7), '1')
Eventos_Excluidos_Porcent, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 8), '1')
Eventos_En_Espera_Cantidad, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 9), '1')
Eventos_En_Espera_Porcent, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 10), '1')
Eventos_Listo_Cantidad, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 11), '1')
Eventos_Listo_Porcent, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 12), '1')
Eventos_en_Error_Cantidad, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 13), '1')
Eventos_en_Error_Porcent, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 14), '1')
Eventos_Perdidos_Cantidad, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 15), '1')
Eventos_Perdidos_Porcent, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 16), '1')
Eventos_Totales_Cantidad, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 17), '1')
Eventos_Totales_Porcent, datos_ultima_compensacion = extractor(datos_ultima_compensacion, chr(fila_index + 18), '1')

print(Tarjetas_Excluidas_Acumuladas)
print(Tarjetas_en_Espera_Acumuladas)
print(Tarjetas_Compensadas)
# print(Eventos_Compensados)
# print(Pago_TDP_a_EPAS_G)

excelObject.save(excelName)

#print("\n" + result2.text)

