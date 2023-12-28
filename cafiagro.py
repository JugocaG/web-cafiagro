from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Ruta del archivo Excel
excel_file_path = 'Datos.xlsx'



# URL del sitio web
url = 'https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094|26566160&rk=813013472&token=b0d99a52-f0e4-4c4a-baba-0ec91ed30b00'

# Función para leer datos de Excel
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



    elemento = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'Invoice'))
     )
    driver.execute_script("arguments[0].click();", elemento)

    elemento_1 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'RadianContributorFileType'))
    )
    driver.execute_script("arguments[0].click();", elemento_1)




    # Iterar sobre los datos y rellenar el formulario


    #     # Rellenar los campos con datos de Excel
    #     campo_nombre.send_keys(nombre)
        
    #     input("Presiona Enter para salir...")

    #     # Puedes agregar más lógica según tus necesidades, como hacer clic en un botón de envío.

finally:
    # Cerrar el navegador al finalizar
    input("Presiona Enter para salir...")
    driver.quit()
