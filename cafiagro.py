from selenium.webdriver.common.by import By
import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Ruta del archivo Excel
excel_file_path = 'Datos.xlsx'

# URL del sitio web
url = 'https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094|26566160&rk=813013472&token=2fa8e527-ea8a-449b-b9dc-0b2344c618db'


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
        EC.visibility_of_element_located((By.ID, 'legalRepresentative'))
    )
    driver.execute_script("arguments[0].click();", elemento)

    NIT_Representante_Legal = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'UserCode'))
    )
    NIT_Representante_Legal.send_keys('26566160')

    NIT_Empresa = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'CompanyCode'))
    )
    NIT_Empresa.send_keys('813013472')

    elemento1 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, 'btn-primary'))
    )
    driver.execute_script("arguments[0].click();", elemento1)


try:

    driver.get(url)

    try:
        iniciar_sesion()

    finally:
        elemento2 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'Invoice'))
        )
        driver.execute_script("arguments[0].click();", elemento2)

        url1 = 'https://catalogo-vpfe.dian.gov.co/User/RedirectToBiller'
        driver.get(url1)

        elemento3 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'menu-button documento'))
        )
        driver.execute_script("arguments[0].click();", elemento3)

        datos_excel = leer_datos_desde_excel(excel_file_path)




finally:
    # Cerrar el navegador al finalizar
    input("Presiona Enter para salir...")
    driver.quit()
