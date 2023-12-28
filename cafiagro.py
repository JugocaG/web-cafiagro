from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import openpyxl
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

def iniciar_sesion():
    elemento = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, 'legalRepresentative'))
    )   

    # Hacer clic en el elemento utilizando JavaScriptExecutor
    driver.execute_script("arguments[0].click();", elemento)

    NIT_Representante_Legal = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'UserCode'))
    )

    # Introducir texto en el elemento de entrada
    NIT_Representante_Legal.send_keys('26566160')

    NIT_Empresa = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'CompanyCode'))
    )

    # Introducir texto en el elemento de entrada
    NIT_Empresa.send_keys('813013472')

    elemento = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CLASS_NAME, 'btn-primary'))
    )   

    # Hacer clic en el elemento utilizando JavaScriptExecutor
    driver.execute_script("arguments[0].click();", elemento)


try:

    driver.get(url)

    datos_excel = leer_datos_desde_excel(excel_file_path)
    try:
        iniciar_sesion()
    finally:
        print("Hola que hace")



finally:
    # Cerrar el navegador al finalizar
    input("Presiona Enter para salir...")
    driver.quit()
