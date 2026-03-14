from docxtpl import DocxTemplate
from docx import Document
from datetime import datetime
import calendar
import pandas as po


df = po.read_excel("CTDR_Asistente_Desarrollo_Informatico.xlsx", sheet_name="TDR Asistente v02", engine = "openpyxl")
actividad1 = df.iloc[33,1]
producto1 = df.iloc[33,6]
producto2 = df.iloc[34,6]
proyecto = df.iloc[9,4].upper()
puesto = df.iloc[10,4]


fecha_actual = datetime.now()

inicio_periodo = fecha_actual.replace(day=1).strftime("%d-%m-%Y")

ultimo_dia = calendar.monthrange(fecha_actual.year, fecha_actual.month)[1]

fin_periodo = fecha_actual.replace(day=ultimo_dia).strftime("%d-%m-%Y")

meses = {
    1: "enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo", 6:"junio", 7:"julio", 8:"agosto", 9:"septiembre", 10:"octubre", 11:"noviembre", 12:"diciembre"
}

mes_Actual=meses[fecha_actual.month]

# Seccion de Fechas para el periodo escrito en texto natural

hoy = datetime.now()
inicio = hoy.replace(day=1)
ultimo = calendar.monthrange(hoy.year, hoy.month)[1]
fin = hoy.replace(day = ultimo)

periodo_incluir_info = f"{inicio.day:02d} al {fin.day:02d} de {meses[hoy.month]} de {hoy.year}"


print("Revisar archivo.")


plantilla_info_productos = DocxTemplate("PLANTILLA_INFORME_PRODUCTOS_JINJA2.docx")
context_info_prod = {
    "numero": "1", # Depende del informe a automatizar
    "producto": producto1.upper(),
    "proyecto":proyecto,
    "mes":mes_Actual.upper(), 
    "funcionario":"BRYAN BENJAMÍN SARABINO CUICHÁN",
    "cedula":"172231964-5",
    "puesto":puesto,
    "fecha_inicio": inicio_periodo,
    "fecha_fin": fin_periodo,
    "periodo": periodo_incluir_info,
    "actividad_1": actividad1,
    "producto_1":producto1,
    "conclusion_1":"Conclusion a redactar 1",
    "conclusion_2":"Conclusion a redactar 2",
    "conclusion_3":"Conclusion a redactar 3",
    "recomendacion_1":"Recomendacion a redactar 1",
    "recomendacion_2":"Recomendacion a redactar 2",
    "recomendacion_3":"Recomendacion a redactar 3",
    "anexo_1": producto1,
    "anexo_2":producto2

}
plantilla_info_productos.render(context_info_prod)

plantilla_info_productos.save("PRUEBA_AUTOMATIZACION_PRODUCTO_1.docx")
