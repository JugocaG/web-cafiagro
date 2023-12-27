from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import openpyxl
import time
from selenium import webdriver

# Ruta del archivo Excel
excel_file_path = 'Datos.xlsx'


# URL del sitio web
url = 'https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094|26566160&rk=813013472&token=d28b82a6-688f-4976-bf8f-41c9f23198ea'

# Función para leer datos de Exce
def leer_datos_desde_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(row)

    return data

# Configurar Selenium y abrir el sitio web
driver = webdriver.Chrome()



try:

    driver.get(url)

    # Abrir otra URL en la misma pestaña
    # Abrir una nueva pestaña en blanco

    # Realizar acciones en la segunda pestaña
    # Puedes interactuar con elementos en la segunda pestaña aquí
    # Leer datos desde Excel
    datos_excel = leer_datos_desde_excel(excel_file_path)
    time.sleep(5)
    boton = driver.find_element(By.ID,'RadianContributorFileType')
    time.sleep(5)
    boton.click()
    # Iterar sobre los datos y rellenar el formulario


    #     # Rellenar los campos con datos de Excel
    #     campo_nombre.send_keys(nombre)
        
    #     input("Presiona Enter para salir...")

    #     # Puedes agregar más lógica según tus necesidades, como hacer clic en un botón de envío.

finally:
    # Cerrar el navegador al finalizar
    input("Presiona Enter para salir...")
    driver.quit()
