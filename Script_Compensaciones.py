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

# Posicion ultima compensacion segun fecha actual (referencia)
posicion_inicio = content.find(fecha_referencia)
datos_ultima_compensacion = content[posicion_inicio:].replace("</pre>", "")

# Iniciamos un indice para barrido de fechas
i = 1

# Bucle que obtiene fecha y datos de la compensacion mas reciente, 
# barriendo desde hoy a dias anteriores
while datos_ultima_compensacion == ">": # Se compara con valor de datos_ultima_compensacion cuando no hubo compensaciones en la fecha de hoy
    fecha_compensacion = date.today() - timedelta(days = i)
    fecha_compensacion = fecha_compensacion.strftime('%Y-%m-%d')
    posicion_inicio = content.find(fecha_compensacion)
    datos_ultima_compensacion = content[posicion_inicio:].replace("</pre>", "")
    i = i + 1

print(datos_ultima_compensacion)

##########################################################################################
###################################### MODULO EXCEL ######################################
##########################################################################################

# Dev
# excelName = datetime.today()
# excelName = excelName.strftime('%Y-%m-%d %H_%M_%S.xlsx')
# excelObject = Workbook()
# excelObject.save(excelName)
# hojaName = 'Hojarda'
# hojaObject = excelObject.create_sheet(hojaName)
# excelObject.save(excelName)

# Casi Prod
excelName = "Planilla de Tiempos de Compensaciones de PY.xlsx"
hojaName = "Datos Backend"
excelObject = load_workbook(excelName)
hojaObject = excelObject[hojaName]

# Dev
# hojaObject = excelObject[hojaName]
# value = hojaObject['B21'].value #5674
# print(value)

fecha_index = None

# Ciclo para obtener el indice de la celda donde esta la ultima compensacion
index_tabla = 7 # Un valor de referencia para reducir la cantidad de iteraciones
while fecha_index != fecha_compensacion:
    index_tabla = index_tabla + 1
    fecha_index = hojaObject['A{index_tabla}'.format(index_tabla = index_tabla)].value.strftime('%Y-%m-%d')

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
# hojaObject['B22'] = "Texto Prueba"




# excelObject.save(excelName)



#wb.save(excelname)



# ################# Variables #################
# extracted_text = datos_ultima_compensacion[:datos_ultima_compensacion.find(";")]
# datos_ultima_compensacion = datos_ultima_compensacion[datos_ultima_compensacion.find(";") + 1:]

def extractor_datos(texto):
    textoExtraido = texto[:texto.find(";")]
    texto = texto[texto.find(";") + 1:]

    # hojaObject['{fila}{columna}'.format(fila = fila, columna = columna)] = textoExtraido
 
    return textoExtraido, texto

def rellena_celda(texto, fila, columna):
    hojaObject['{fila}{columna}'.format(fila = fila, columna = columna)] = texto

fila_index = 65 # 'A'
columna_index = index_tabla

Fecha_base, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Tarjetas_Excluidas_Acumuladas, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Tarjetas_en_Espera_Acumuladas, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Tarjetas_Compensadas, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_Compensados, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Pago_TDP_a_EPAS_GS, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Pago_TDP_a_EPAS_USD, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_Excluidos_Cantidad, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_Excluidos_Porcent, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_En_Espera_Cantidad, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_En_Espera_Porcent, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_Listo_Cantidad, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_Listo_Porcent, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_en_Error_Cantidad, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_en_Error_Porcent, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_Perdidos_Cantidad, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_Perdidos_Porcent, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_Totales_Cantidad, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)
Eventos_Totales_Porcent, datos_ultima_compensacion = extractor_datos(datos_ultima_compensacion)

# rellena_celda(Fecha_base, chr(fila_index + 0), columna_index)
rellena_celda(Tarjetas_Excluidas_Acumuladas, chr(fila_index + 1), columna_index)
rellena_celda(Tarjetas_en_Espera_Acumuladas, chr(fila_index + 2), columna_index)
rellena_celda(Tarjetas_Compensadas, chr(fila_index + 3), columna_index)
rellena_celda(Eventos_Compensados, chr(fila_index + 4), columna_index)
rellena_celda(Pago_TDP_a_EPAS_GS, chr(fila_index + 5), columna_index)
rellena_celda(Pago_TDP_a_EPAS_USD, chr(fila_index + 6), columna_index)
rellena_celda(Eventos_Excluidos_Cantidad, chr(fila_index + 7), columna_index)
rellena_celda(Eventos_Excluidos_Porcent, chr(fila_index + 8), columna_index)
rellena_celda(Eventos_En_Espera_Cantidad, chr(fila_index + 9), columna_index)
rellena_celda(Eventos_En_Espera_Porcent, chr(fila_index + 10), columna_index)
rellena_celda(Eventos_Listo_Cantidad, chr(fila_index + 11), columna_index)
rellena_celda(Eventos_Listo_Porcent, chr(fila_index + 12), columna_index)
rellena_celda(Eventos_en_Error_Cantidad, chr(fila_index + 13), columna_index)
rellena_celda(Eventos_en_Error_Porcent, chr(fila_index + 14), columna_index)
rellena_celda(Eventos_Perdidos_Cantidad, chr(fila_index + 15), columna_index)
rellena_celda(Eventos_Perdidos_Porcent, chr(fila_index + 16), columna_index)
rellena_celda(Eventos_Totales_Cantidad, chr(fila_index + 17), columna_index)
rellena_celda(Eventos_Totales_Porcent, chr(fila_index + 18), columna_index)




print(Tarjetas_Excluidas_Acumuladas)
print(Tarjetas_en_Espera_Acumuladas)
print(Tarjetas_Compensadas)
# print(Eventos_Compensados)
# print(Pago_TDP_a_EPAS_G)

excelObject.save(excelName)

#print("\n" + result2.text)

