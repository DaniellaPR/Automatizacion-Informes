from docxtpl import DocxTemplate
from docx import Document
from datetime import datetime
import calendar
import pandas as po


df = po.read_excel("CTDR_Asistente_Desarrollo_Informatico.xlsx", sheet_name="TDR Asistente v02", engine = "openpyxl")

# Se determina cuantos productos existen en el TDR
numero_inicial_prod = df.iloc[33:,0]

numeros = numero_inicial_prod[numero_inicial_prod.apply(lambda x: isinstance(x, (int)))]
    
ultimo_valor = numeros.iloc[-1] if not numeros.empty else None


# Con el último valor es posible recorrer exactamente 
# las actividades y productos para almacenarlos en listas

# Posición Inicial Actividades
columna_actividades = 1
fila_actividades = 33 # Recordemos que usamos ultimo_valor para saber cuantas posiciones recorrer

# Extraen valores de actividad
serie_actividades = df.iloc[fila_actividades:fila_actividades + ultimo_valor, columna_actividades]
actividades_lista = serie_actividades.tolist()
print(actividades_lista)

# Posicion Inicial de los Productos
columna_productos = 6
fila_productos = 33 # Recordemos que usamos ultimo_

# Extraen valores de los Productos
serie_productos = df.iloc[fila_productos: fila_productos + ultimo_valor, columna_productos]
productos_lista = serie_productos.tolist()
print(productos_lista)

# Una vez creadas las listas, se procede a extraer información importante para el contexto generalizado

# Se extraen los valores del proyecto, puesto y honorarios
proyecto = df.iloc[9,4].upper()
puesto = df.iloc[10,4]
honorario = df.iloc[68,4]


# Se extraen valores de fechas

fecha_actual = datetime.now()

inicio_periodo = fecha_actual.replace(day=1).strftime("%d-%m-%Y")

ultimo_dia = calendar.monthrange(fecha_actual.year, fecha_actual.month)[1]

fin_periodo = fecha_actual.replace(day=ultimo_dia).strftime("%d-%m-%Y")

meses = {
    1: "enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo", 6:"junio", 7:"julio", 8:"agosto", 9:"septiembre", 10:"octubre", 11:"noviembre", 12:"diciembre"
}

mes_Actual=meses[fecha_actual.month]

##### Plazos extracción ####
metodologia = df.iloc[23,0]
lst_metodologia = metodologia.split()
indices = [i for i, palabra in enumerate(lst_metodologia)
           if palabra =="Fecha:"]
plazos = []
for i in indices:
    rango = lst_metodologia[i+2:i+10]
    plazos.append(" ".join(rango))
if len(indices)>2:
    segunda = lst_metodologia[indices[1]+2]
    print(" ".join(segunda))
plazo_1=plazos[0]
plazo_2=plazos[1]


# Seccion de Fechas para el periodo escrito en texto natural

hoy = datetime.now()
inicio = hoy.replace(day=1)
ultimo = calendar.monthrange(hoy.year, hoy.month)[1]
fin = hoy.replace(day = ultimo)

periodo_incluir_info = f"{inicio.day:02d} al {fin.day:02d} de {meses[hoy.month]} de {hoy.year}"

# Se tiene que idear una forma dinámica de llenar el campo de producto y actividades dependiendo de cuantas actividades o productos existan
# Pero por lo pronto se llenan los otros campos con lo que ya tenemos


# Se crea el contexto para utilizarlo en la herramienta Docx
contexto_plantilla = {
    #"numero": ultimo_valor,
    #"producto": productos_lista, #Especial ciudado con este campo
    "proyecto": proyecto,
    "mes": mes_Actual.upper(),
    "funcionario": "BRYAN BENJAMÍN SARABINO CUICHÁN",
    "cedula": "172231964-5",
    "puesto": puesto,
    "honorario":honorario,
    "fecha_inicio":inicio_periodo,
    "fecha_fin":fin_periodo,
    "periodo":periodo_incluir_info,
    #"actividad":actividades_lista, #Especial cuidado con este campo
    "conclusion_1":"Escribir conclusión #1",
    "conclusion_2":"Escribir conclusión #2",
    "conclusion_3":"Escribir conclusión #3",
    "recomendacion_1":"Escribir recomendación #1",
    "recomendacion_2":"Escribir recomendación #2",
    "recomendacion_3":"Escribir recomendación #3"

}

# Una vez hemos creado el contexto general para la 
# plantilla de productos, se tiene que proceder 
# a llenar los campos


##############################################################################################

# Sección de Pruebas

# Se abre el archivo de excel con el path para las plantillas
df_pl = po.read_excel("PLANTILLAS_INFORMES.xlsx", sheet_name="Hoja1", engine = "openpyxl")

# Ahora se realiza una prueba con un menu simple (FUNCIONAL PARA LOS PRODUCTOS)

print("1. Informe de Productos")
print("2. Informe de Actividades y Productos Entregados")
print("3. Informe de Aceptación de los Productos Recibidos a Satisfacción")

seleccion_menu = input("Indique el informe (1 - 3): ")

if (seleccion_menu == "1"):
    print()
    print("Elija cual informe realizar")
    for i in range(ultimo_valor):
        print(f"{i + 1}. Producto {i + 1}")
        print()
    print(f"{ultimo_valor + 1}. Generar todos los Informes")

    usuario_seleccion = input(f"Digite su elección (1 - {ultimo_valor}): ")

    if(int(usuario_seleccion) <= ultimo_valor):
        df_pl = po.read_excel("PLANTILLAS_INFORMES.xlsx", sheet_name="Hoja1", engine = "openpyxl")
        
        path_informe = df_pl.iloc[0,2]

        plantilla_prod = DocxTemplate(path_informe)

        contexto_actividad_prod = {
            **contexto_plantilla,
            "producto":productos_lista[int(usuario_seleccion) - 1],
            "actividad":actividades_lista[int(usuario_seleccion) - 1],
            "numero": usuario_seleccion

        }
        plantilla_prod.render(contexto_actividad_prod)
        plantilla_prod.save(f"Informe_de_Producto_{usuario_seleccion}.docx")

        print("Revise su directorio.")
        print()
    elif (int(usuario_seleccion) == (ultimo_valor + 1)):
        df_pl = po.read_excel("PLANTILLAS_INFORMES.xlsx", sheet_name="Hoja1", engine = "openpyxl")
        path_informe = df_pl.iloc[0,2]

        for i in range(1,ultimo_valor + 1):
            plantilla_prod = DocxTemplate(path_informe)
            contexto_actividad_prod = {
                **contexto_plantilla,
                "producto":productos_lista[i - 1],
                "actividad":actividades_lista[i - 1],
                "numero": i
            }
            plantilla_prod.render(contexto_actividad_prod)
            plantilla_prod.save(f"Informe_de_Producto_{i}.docx")
        print("Revise su directorio.")
        print()

    else:
        print("Error.")
elif (seleccion_menu == "2"):

    print()
    print("Indique cuantos productos y/o actividades tiene el informe:")
    for i in range(ultimo_valor):
        print(f"{i + 1}. Producto / Actividad {i + 1}")
        print()
    print(f"{ultimo_valor + 1}. Actividades y Productos Completados")
    usuario_seleccion = input(f"Digite su elección (1 - {ultimo_valor}): ")

    if(int(usuario_seleccion) <= ultimo_valor):
        df_pl = po.read_excel("PLANTILLAS_INFORMES.xlsx", sheet_name="Hoja1", engine = "openpyxl")
        
        path_informe = df_pl.iloc[3,2]

        plantilla_prod_act = DocxTemplate(path_informe)

        contexto_actividad_prod = {
            **contexto_plantilla,
            "producto":productos_lista[int(usuario_seleccion) - 1],
            "actividad":actividades_lista[int(usuario_seleccion) - 1],
            "numero": usuario_seleccion

        }
        plantilla_prod_act.render(contexto_actividad_prod)
        plantilla_prod_act.save(f"Informe_de_Actividad_Producto_Realizados_{usuario_seleccion}.docx")

        print("Revise su directorio.")
        print()

    elif (int(usuario_seleccion) == (ultimo_valor + 1)):
        df_pl = po.read_excel("PLANTILLAS_INFORMES.xlsx", sheet_name="Hoja1", engine = "openpyxl")
        path_informe = df_pl.iloc[1,2]

        plantilla_prod_act = DocxTemplate(path_informe)

        contexto_actividad_prod = {
            **contexto_plantilla,
            "productos_lista":productos_lista,
            "actividades_lista":actividades_lista,
        }
        plantilla_prod_act.render(contexto_actividad_prod)
        plantilla_prod_act.save("Informe_de_Actividad_Producto_Realizados_Total.docx")

        print("Revise su directorio")
        print()
    else:
        print("Error.")
        print()

elif (seleccion_menu == "3"):
    print(plazos)
    

"""
# Contexto para rellenar plantilla
context_info = {
    "numero": ultimo_valor, 
    "producto": producto1.upper(),
    "proyecto":proyecto,
    "mes":mes_Actual.upper(), 
    "funcionario":"BRYAN BENJAMÍN SARABINO CUICHÁN",
    "cedula":"172231964-5",
    "puesto":puesto,
    "honorario": honorario,
    "plazo_1": plazo_1,
    "plazo_2": plazo_2,
    "fecha_inicio": inicio_periodo,
    "fecha_fin": fin_periodo,
    "periodo": periodo_incluir_info,
    "actividad_1": actividad1,
    "producto_1":producto1.upper(),
    "actividad_2": actividad2,
    "producto_2": producto2,
    "conclusion_1":"Conclusion a redactar 1",
    "conclusion_2":"Conclusion a redactar 2",
    "conclusion_3":"Conclusion a redactar 3",
    "recomendacion_1":"Recomendacion a redactar 1",
    "recomendacion_2":"Recomendacion a redactar 2",
    "recomendacion_3":"Recomendacion a redactar 3",
    "anexo_1": producto1,
    "anexo_2":producto2

}

"""
