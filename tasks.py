from robocorp.tasks import task
from RPA.Assistant import Assistant
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import matplotlib.pyplot as plt
import pandas as pd
from RPA.Excel.Application import Application
import speech_recognition as sr
import pyaudio
@task
def first_view():
    assistant = Assistant()
    assistant.add_heading("Pacientes")
    assistant.add_text_input("text_input", placeholder="CODIGO DE ALUMNO")
    assistant.add_submit_buttons("Submit", default="Submit")
    assistant.add_heading("Informe del Dia")
    assistant.add_button("Generar", generar_reporte_pdf)
    result = assistant.run_dialog()
    codigo = result.text_input
    verificar(codigo)

def leer_datos_excel():
    excel = Files()
    excel.open_workbook("Libro1.xlsx")
    datos = excel.read_worksheet_as_table("FICHA", header=True)
    excel.close_workbook()
    return datos


def verificar(alumno1):
    registro=buscar_excel(alumno1)
    if registro!='ninguno':
        mostrar_datos_alumno(registro)
    else:
        first_view()

def generar_graficabarras1():
    datos=leer_datos_excel()
    df = pd.DataFrame(datos)
    frequencies = df['FACULTAD'].value_counts()

    frequencies.plot(kind='bar')
    plt.title('Frecuencia de Facultades')
    plt.xlabel('FACULTAD')
    plt.ylabel('Frecuencia')
    plt.savefig("facultad.png")
    plt.close()


def generar_graficabarras2():
    datos=leer_datos_excel()
    df = pd.DataFrame(datos)
    frequencies = df['AREA'].value_counts()

    frequencies.plot(kind='bar')
    plt.title('Frecuencia de Areas Medicas')
    plt.xlabel('FACULTAD')
    plt.ylabel('Frecuencia')
    plt.savefig("Area.png")
    plt.close()

def generar_graficabarras3():
    datos=leer_datos_excel()
    df = pd.DataFrame(datos)
    frequencies = df['ASISTENCIA'].value_counts()

    frequencies.plot(kind='bar')
    plt.title('ASISTENCIAS VS FALTAS')
    plt.ylabel('Frecuencia')
    plt.savefig("Asistencia.png")
    plt.close()



def generar_reporte_pdf():
   
    generar_graficabarras1()
    generar_graficabarras2()
    generar_graficabarras3()

    
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Reporte de Frecuencias y Asistencias</title>
    </head>
    <body>
        <h1>Reporte de Frecuencias y Asistencias</h1>
        <h2>Frecuencia de Facultades</h2>
        <img width="400"  src="facultad.png" alt="Frecuencia de Facultades">
        <h2>Frecuencia de Areas Medicas</h2>
        <img width="500"  src="Area.png" alt="Frecuencia de Areas Medicas">
        <h2>---------------------------</h2>
        <h2>Asistencias vs Faltas</h2>
        <img width="500"  src="Asistencia.png" alt="Asistencias vs Faltas">
    </body>
    </html>
    """

    pdf = PDF()
    pdf.html_to_pdf(html_content, "reportev1.pdf")
    exit()

def mostrar_datos_alumno(paciente):
    assistant = Assistant()
    assistant.add_text("Información del Alumno:")

    for key, value in paciente.items():
        assistant.open_row()
        assistant.add_text(key)
        assistant.add_text(value)
        assistant.close_row()
    assistant.run_dialog()
    rellenar_campos(paciente)

def rellenar_campos(paciente):
    assistant=Assistant()
    assistant.add_text_input("CIEV10", placeholder="CIEV10")
    assistant.add_submit_buttons("Submit", default="Submit")
    result = assistant.run_dialog()
    descr = oir()
    ciev10 = result.CIEV10
    actualizardatos(paciente,ciev10,descr)


def actualizardatos(paciente,ciev10,descr):
    indice=int(paciente["ID"])+1
    app = Application()
    app.open_application()
    app.open_workbook('Libro1.xlsx')
    app.set_active_worksheet(sheetname='FICHA')
    app.write_to_cells(row=indice, column=9, value='Asistencia')
    app.write_to_cells(row=indice, column=10, value=ciev10)
    app.write_to_cells(row=indice, column=11, value=descr)
    app.save_excel()
    app.quit_application()
    paciente["CIEV10"]=ciev10
    paciente["DESCRIPCION"]=descr
    mostrar_datos_alumnonuevo(paciente)

def mostrar_datos_alumnonuevo(paciente):
    assistant = Assistant()
    assistant.add_text("Información del Alumno:")

    for key, value in paciente.items():
        assistant.open_row()
        assistant.add_text(key)
        assistant.add_text(value)
        assistant.close_row()
    assistant.run_dialog()
    first_view()

def buscar_excel(codigo):
    excel = Files()
    excel.open_workbook("Libro1.xlsx")
    worksheet = excel.read_worksheet_as_table("FICHA", header=True)
    excel.close_workbook()
    for row in worksheet:
        if row['COD_ALUMNO'] == codigo:
            return row
    return "ninguno"

def oir():
    with sr.Microphone() as source:
        reconocedor=sr.Recognizer()
        print("habla cualquier cosa..")
        audio=reconocedor.listen(source)
        try:
            text=reconocedor.recognize_google(audio,language='es-ES')
            print(text)
            return text
        except:
            print("intentalo de nuevo")