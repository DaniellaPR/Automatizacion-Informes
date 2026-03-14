from docxtpl import DocxTemplate
from docx import Document
from datetime import datetime
import calendar
import pandas as po


df = po.read_excel("CTDR_Asistente_Desarrollo_Informatico.xlsx", sheet_name="TDR Asistente v02", engine = "openpyxl")
actividad1 = df.iloc[33,1]
actividad2 = df.iloc[34,1]
producto1 = df.iloc[33,6]
producto2 = df.iloc[34,6]
proyecto = df.iloc[9,4].upper()
puesto = df.iloc[10,4]
honorario = df.iloc[68,4]


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

# CONTEXTO PARA RELLENAR PLANTILLAS

# plantilla_info_productos = DocxTemplate("PLANTILLA_INFORME_PRODUCTOS_JINJA2.docx")
context_info = {
    "numero": "1", 
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


# SECCION DE SELECCION DEPENDIENDO DEL TIPO DE INFORME A GENERAR

print("Seleccione el informe que desee generar:")

print("1. Informe de Productos")
print("2. Informe de Actividades y Productos Entregados")
print("3. Informe de Aceptación de los Productos Recibidos a Satisfacción")

decision_usuario = input("Ingrese el número del informe a generar: ")

df_pl = po.read_excel("PLANTILLAS_INFORMES.xlsx", sheet_name="Hoja1", engine = "openpyxl")

if (decision_usuario == "1"):
    print("¿De que producto quiere generar el informe?")
    print("1. Producto N°1")
    print("2. Producto N°2")
    print("3. Todos los productos")


    productos_generar = input("Ingrese su elección: ")


    if (productos_generar == "1"):

        path_informe = df_pl.iloc[0,2]

        plantilla_prod_1 = DocxTemplate(path_informe)

        plantilla_prod_1.render(context_info)

        plantilla_prod_1.save("Informe_de_Productos_Producto_1.docx")

        print("Revise su directorio.")

    elif (productos_generar == "2"):
        path_informe = df_pl.iloc[1,2]

        plantilla_prod_2 = DocxTemplate(path_informe)

        plantilla_prod_2.render(context_info)

        plantilla_prod_2.save("Informe_de_Productos_Producto_2.docx")

        print("Revise su directorio.")

    elif (productos_generar == "3"):
        path_informe_1 = df_pl.iloc[0,2]
        path_informe_2 = df_pl.iloc[1,2]

        plantilla_prod_1 = DocxTemplate(path_informe_1)

        plantilla_prod_1.render(context_info)

        plantilla_prod_1.save("Informe_de_Productos_Producto_1.docx")

        plantilla_prod_2 = DocxTemplate(path_informe_2)

        plantilla_prod_2.render(context_info)

        plantilla_prod_2.save("Informe_de_Productos_Producto_2.docx")

        print("Revise su directorio.")

elif (decision_usuario == "2"):
    path_informe = df_pl.iloc[2,2]

    plantilla_act_prod = DocxTemplate(path_informe)

    plantilla_act_prod.render(context_info)

    plantilla_act_prod.save("Informe_Actividades_Productos.docx")
    
    print("Revise su directorio.")

elif (decision_usuario == "3"):
    path_informe = df_pl.iloc[3,2]

    plantilla_act_prod = DocxTemplate(path_informe)

    plantilla_act_prod.render(context_info)

    plantilla_act_prod.save("Informe_Aceptacion.docx")
    
    print("Revise su directorio.")
