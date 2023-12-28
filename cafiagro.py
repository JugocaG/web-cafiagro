from selenium.webdriver.common.by import By
import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

# Ruta del archivo Excel
excel_file_path = 'Datos.xlsx'

# URL del sitio web
url = 'https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094|26566160&rk=813013472&token=c6a58a3d-db63-4971-acff-11fbc9f68f09'


# Función para leer datos de Excel
def leer_datos_desde_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(row)
    return data

def ingresa_data(data):
    # Configurar Selenium
    driver = webdriver.Chrome()
    driver.get('https://www.facebook.com/')
    # Utilizar los datos en Selenium
    for row in data:
        # Aquí puedes realizar interacciones con Selenium usando los datos
        # Por ejemplo, enviar datos a un formulario, hacer clic en botones, etc.
        input_element = driver.find_element_by_id('email')
        input_element.send_keys(row[0])

    # Realizar más acciones con Selenium según sea necesario

driver = webdriver.Chrome()
try:
    datos_excel = leer_datos_desde_excel(excel_file_path)
    ingresa_data(datos_excel)
finally:
    # Cerrar el navegador al finalizar
    input(":(((")
# Configurar Selenium y abrir el sitio web


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
            EC.element_to_be_clickable((By.CLASS_NAME, 'menu-button.documento'))
        )
        driver.execute_script("arguments[0].click();", elemento3)




        dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'RangoNum'))
        )

        # Crear un objeto Select para interactuar con el elemento de lista desplegable
        select = Select(dropdown)

        # Seleccionar una opción por valor
        select.select_by_value('a1622672-35d6-4132-a6a2-491657083a98')

        elemento_vendedor = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseTwo"]'))
        )

        # Hacer clic en el elemento
        elemento_vendedor.click()

        

        codigo_postal = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'accountingCustomerPartyField_Party_PhysicalLocation_Address_PostalZone'))
        )

        select_codigo_postal = Select(codigo_postal)

        # Seleccionar una opción por valor
        select_codigo_postal.select_by_value('111711')



        id_del_select = 'origin'

        # Nuevo valor que deseas establecer
        nuevo_valor = '10'

        # Ejecutar JavaScript para cambiar el valor del elemento <select>
        script = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script)

        

    


        
        




finally:
    # Cerrar el navegador al finalizar
    input("Presiona Enter para salir...")
    driver.quit()
