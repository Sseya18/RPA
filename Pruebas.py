from robocorp.tasks import task
from RPA.Assistant import Assistant
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import matplotlib.pyplot as plt
import pandas as pd

@task
def first_view():
    assistant = Assistant()
    assistant.add_heading("Pacientes")
    assistant.add_text_input("text_input", placeholder="CODIGO DE ALUMNO")
    assistant.add_submit_buttons("Submit", default="Submit")
    assistant.add_heading("Informe del Dia")
    assistant.add_button("Generar", generar_graficas)
    result = assistant.run_dialog()
    codigo = result.text_input
    verificar(codigo)

def verificar(codigo):
    # Implementa tu lógica de verificación aquí
    print(f"Código ingresado: {codigo}")

def leer_datos_excel():
    excel = Files()
    excel.open_workbook("Libro1.xlsx")
    datos = excel.read_worksheet_as_table("FICHA", header=True)
    excel.close_workbook()
    return datos

def generar_graficas():
    datos = leer_datos_excel()
    generar_grafica(datos)

def generar_grafica(datos):
    df = pd.DataFrame(datos)

    # Contar la frecuencia de cada valor en la columna 'Área'
    frecuencia = df['Área'].value_counts()

    # Crear la gráfica de barras
    frecuencia.plot(kind='bar')
    plt.title('Frecuencia de Área')
    plt.xlabel('Área')
    plt.ylabel('Frecuencia')
    plt.savefig("grafica.png")
    plt.close()