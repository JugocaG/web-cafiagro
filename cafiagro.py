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
url = 'https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094|26566160&rk=813013472&token=7bcaa1a6-8118-40a8-a63d-a4d6995ef3d4'


# Función para leer datos de Excel
def leer_datos_desde_excel():
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active
    data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(row)
    print(data)
    return data

leer_datos_desde_excel()

def ingresa_data(data):
    # Configurar Selenium
    driver = webdriver.Chrome()
    driver.get('https://accounts.google.com/lifecycle/steps/signup/name?continue=https://myaccount.google.com?utm_source%3Daccount-marketing-page%26utm_medium%3Dcreate-account-button&dsh=S805779475:1703802835653781&flowEntry=SignUp&flowName=GlifWebSignIn&theme=glif&TL=AHNYTIRXqx6ycBtWEQriyBkjJw-pE9xfZJIzxEUeGhdRMRiJBVhUWVbGayK0a8vn')
    # Utilizar los datos en Selenium

    try:
        for row in data:
            input_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'firstName'))
            )
            input_element.send_keys(row[0])  # row[0] es el primer elemento en la fila actual

            input_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'lastName'))
            )
            input_element.send_keys(row[1])  # row[1] es el segundo elemento en la fila actual

            # Puedes agregar más campos y datos según sea necesario

            # Esperar un breve momento antes de pasar a la siguiente iteración (opcional)
            driver.implicitly_wait(5)
        # input_element = WebDriverWait(driver, 10).until(
        #     EC.presence_of_element_located((By.ID, 'firstName'))
        # )
        # input_element.send_keys(
        #     data[0][0])  # data[0] representa la primera fila, data[0][0] es el primer elemento en esa fila
        #
        # input_element = WebDriverWait(driver, 10).until(
        #     EC.presence_of_element_located((By.ID, 'lastName'))
        # )
        # input_element.send_keys(data[0][1])
        # for row in data:
        #     # Aquí puedes realizar interacciones con Selenium usando los datos
        #     # Por ejemplo, enviar datos a un formulario, hacer clic en botones, etc.
        #
        #     input_element = WebDriverWait(driver, 10).until(
        #         EC.presence_of_element_located((By.ID, 'firstName'))
        #     )
        #     input_element.send_keys(row[0])
        #
        #     input_element = WebDriverWait(driver, 10).until(
        #         EC.presence_of_element_located((By.ID, 'lastName'))
        #     )
        #     input_element.send_keys(row[1])
    finally:
        # Cerrar el navegador al finalizar
        input(":)")

try:
    datos_excel = leer_datos_desde_excel()
    ingresa_data(datos_excel)
finally:
     # Cerrar el navegador al finalizar
    input(":)")
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

       # elemento2 = WebDriverWait(driver, 10).until(
        #    EC.element_to_be_clickable((By.CLASS_NAME, 'parent-link'))
        #)
        #driver.execute_script("arguments[0].click();", elemento2)

        #elemento4 = WebDriverWait(driver, 5).until(
        #    EC.element_to_be_clickable((By.ID, 'RadianContributorFileType'))
        #)
        #driver.execute_script("arguments[0].click();", elemento4)

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

        id_del_select1 = 'RegimenFiscal'
         #Nuevo valor que deseas establecer
        nuevo_valor1 = 'R-99-PN'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
        driver.execute_script(script1)

        id_del_select1 = 'TipoDocumento'
         #Nuevo valor que deseas establecer
        nuevo_valor1 = '31'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
        driver.execute_script(script1)

        id_del_select1 = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_Country_IdentificationCode'
         #Nuevo valor que deseas establecer
        nuevo_valor1 = 'CO'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
        driver.execute_script(script1)

        id_del_select1 = 'NumeroDocumento'
         #Nuevo valor que deseas establecer
        nuevo_valor1 = '12143849'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
        driver.execute_script(script1)

        id_del_select1 = 'Departamento'
         #Nuevo valor que deseas establecer
        nuevo_valor1 = '41'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
        driver.execute_script(script1)
        script_change_event = f"document.getElementById('{id_del_select1}').dispatchEvent(new Event('change'));"
        driver.execute_script(script_change_event)








finally:
    # Cerrar el navegador al finalizar
    input("Presiona Enter para salir...")
    driver.quit()
