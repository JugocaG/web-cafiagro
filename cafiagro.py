from selenium.webdriver.common.by import By
import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime
import pytz
import time
import pyautogui

# Ruta del archivo Excel
excel_file_path = 'Datos.xlsx'

# URL del sitio web
url = 'https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094%7C26566160&rk=813013472&token=3e8a56d0-eb02-4bd5-b3f7-4716269250f1'


# Función para leer datos de Excel
def leer_datos_desde_excel():
    workbook = openpyxl.load_workbook(excel_file_path)
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
#try:
    #datos_excel = leer_datos_desde_excel(excel_file_path)
    #ingresa_data(datos_excel)
#finally:
    # Cerrar el navegador al finalizar
    #input(":(((")
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

def llenar_campos():
    #-------------------------------------------------------- DATOS DEL DOCUMENTO --------------------------------------------------------#
    
    dropdown = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, 'RangoNum'))
    )
    select = Select(dropdown)
    select.select_by_value('a1622672-35d6-4132-a6a2-491657083a98')
    #Orden de Compra
    id_del_select = 'OrderReference_ID' 
    nuevo_valor = '4145'
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)
    # ------------------------------ #

    elemento10 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CLASS_NAME, 'ui-datepicker-trigger'))
    )
    elemento10.click()

    elemento20 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CLASS_NAME, 'ui-state-default.ui-state-highlight.ui-state-hover'))
    )
    elemento20.click()
    
    id_del_select = 'paymentMeansField_PaymentMeansCode'
    nuevo_valor = '1'
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)
    id_del_select = 'TiopoNeg'
    nuevo_valor = '1'
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)
    
    #-------------------------------------------------------- DATOS DEL ADQUIRIENTE / COMPRADOR --------------------------------------------------------#

    elemento_datos = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseVendor"]'))
    )
    elemento_datos.click()
    id_del_select = 'accountingCustomerPartyField_Party_PartyTaxScheme_TaxLevelCode'
    nuevo_valor = 'R-99-PN'
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)
    id_del_select = 'ResponsabilidadTributariaAdquiriente'
    nuevo_valor = 'ZZ'
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)

    #-------------------------------------------------------- DATOS DEL VENDEDOR --------------------------------------------------------#

    elemento_vendedor = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseTwo"]'))
    )
    elemento_vendedor.click()

    #Procedencia
    id_del_select = 'origin'
    nuevo_valor = '10'
    script = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script)

    #Tipo de documento
    id_del_select1 = 'TipoDocumento'
    nuevo_valor1 = '31'
    script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
    driver.execute_script(script1)

    #Numero de documento
    id_del_select1 = 'NumeroDocumento'
    nuevo_valor1 = '12181767'
    script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
    driver.execute_script(script1)

    #Razon social
    id_del_select = 'RazonSocial'
    nuevo_valor = 'ISAIAS ISAIAS ISAIAS'
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)

    #Tipo de Contribuyente
    id_del_select = 'tipoContribuyente'
    nuevo_valor = '2'
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)

    #Responsabilidad Tributaria
    id_del_select = 'ResposabilidadTributaria'
    nuevo_valor = 'ZZ'
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)

    #Regimen Fiscal
    id_del_select1 = 'RegimenFiscal'
    nuevo_valor1 = 'R-99-PN'
    script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
    driver.execute_script(script1)

    #Pais
    id_del_select1 = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_Country_IdentificationCode'
    nuevo_valor1 = 'CO'
    script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
    driver.execute_script(script1)

    #Departamento
    id_del_select = 'Departamento'
    nuevo_valor = '41'        
    nuevo_texto = '41 | Huila'
    script = f"var select = document.getElementById('{id_del_select}'); " \
            f"var option = document.createElement('option'); " \
            f"option.value = '{nuevo_valor}'; " \
            f"option.text = '{nuevo_texto}'; " \
            f"select.add(option);"
    driver.execute_script(script)
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)
    
    #Ciudad
    id_del_select = 'Ciudad'
    nuevo_valor = '41668'
    nuevo_texto = '41668 | SAN AGUSTÍN'
    script = f"var select = document.getElementById('{id_del_select}'); " \
            f"var option = document.createElement('option'); " \
            f"option.value = '{nuevo_valor}'; " \
            f"option.text = '{nuevo_texto}'; " \
            f"select.add(option);"
    driver.execute_script(script)
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)
    
    #Direccion
    id_del_select = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_AddressLine'
    nuevo_valor = 'VEREDA LOS SAUCES'
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)

    #Codigo Postal
    id_del_select = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_PostalZone'
    nuevo_valor = '417030'
    nuevo_texto = '417030'
    script = f"var select = document.getElementById('{id_del_select}'); " \
            f"var option = document.createElement('option'); " \
            f"option.value = '{nuevo_valor}'; " \
            f"option.text = '{nuevo_texto}'; " \
            f"select.add(option);"
    driver.execute_script(script)
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)

    #-------------------------------------------------------- DETALLE DE PRODUCTO / SERVICIO --------------------------------------------------------#

    elemento_vendedor = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseTwo"]'))
    )

    # Hacer clic en el elemento
    elemento_vendedor.click()
    time.sleep(0.5)
    elemento_productos = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseProductDetails"]'))
    )
    elemento_productos.click()

    driver.execute_script("uploadListProductDetails(1);")

    elemento_datos = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="IdTableListProduct"]/tbody/tr[1]/td[3]'))
    )

    # Hacer clic en el elemento
    elemento_datos.click()

    #Precio Unitario
    id_del_select1 = 'PrecioUnitario1'
    precio_unitario = '5077.83'
    script1 = f"document.getElementById('{id_del_select1}').value = '{precio_unitario}';"
    driver.execute_script(script1)

    elemento_datos = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="Cantidad1"]'))
    )

    # Hacer clic en el elemento cantidad
    elemento_datos.click()
    time.sleep(1)
    pyautogui.typewrite('626.8')

    elemento_datos = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="ProductDetailsTable"]/tbody/tr/td[4]'))
    )
    elemento_datos.click()

    #Forma de generación y transmisión
    id_del_select = 'Formagen1'
    nuevo_valor = 1
    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
    driver.execute_script(script1)

    #Forma de generación y transmisión
    zona_horaria = pytz.timezone('America/Bogota')
    fecha_actual = datetime.now(zona_horaria)
    fecha_local = fecha_actual.date()
    id_del_select = 'FechaCompra1'
    script1 = f"document.getElementById('{id_del_select}').value = '{fecha_local}';"
    driver.execute_script(script1)

    #-------------------------------------------------------- ABRIR FIRMAR Y GUARDAR --------------------------------------------------------#


    elemento1 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, 'btnPreView'))
    )
    driver.execute_script("arguments[0].click();", elemento1)

    time.sleep(3)

    elemento32 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, 'btnFimarGuardar'))
    )
    driver.execute_script("arguments[0].click();", elemento32)

    time.sleep(3)

    driver.switch_to.alert.accept()

    driver.get('https://gratis-vpfe.dian.gov.co/SupportDocuments/Adjustment')


driver.get(url)
try:
    iniciar_sesion()
finally:

    driver.get('https://catalogo-vpfe.dian.gov.co/User/RedirectToBiller')
    elemento3 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, 'menu-button.documento'))
    )
    driver.execute_script("arguments[0].click();", elemento3)

    llenar_campos()
    
    
            
    
    # Cerrar el navegador al finalizar
    input("Presiona Enter para salir...")
    driver.quit()
